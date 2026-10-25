"""
script.py
=========
Reads student marks from an Excel workbook and inserts answer rows into the
MySQL `answer` table - replacing manual data entry.

FULLY DYNAMIC - nothing about the question count, question types, marks,
competency codes, or sheet names is hardcoded. Everything is detected at run
time from the workbook and the database:

  * Sheets     : auto-detected (any worksheet whose A1="competency" and A2="Students").
  * Section    : derived from the sheet name (leading letters), overridable via --sections.
  * Questions  : count + type + max marks read from examquestion + questions for the paper.
  * Competency : read from Excel row 1 (whatever codes appear there).
  * Students   : scanned from row 3 until the first blank student-id cell.

Answer-fill logic is driven by each question's real `questions.quesType`:
  - multipleChoice : marks>0 -> correct option, else a different option.
  - trueFalse      : marks>0 -> correct value, else the opposite.
  - anything else  : uanswer and canswer stay NULL (open-ended).

Run:
    python script.py --all --overwrite                 # all detected sheets, replace
    python script.py --sheets A1,B1                    # specific sheets only
    python script.py --sheet B1                        # one sheet
    python script.py --paper 90 --class 9 --all        # different exam
    python script.py --all --no-insert --save-sql      # dry run + .sql backup
    python script.py --all --no-insert --save-sql      # dry run + .sql backup
    python script.py --paper 76 --export-meta     # write paper_76_meta.json (for SQL generation in the dashboard)
"""

import argparse
import json
import math
import re
from pathlib import Path

import openpyxl
import mysql.connector

from config import (
    DEFAULT_PAPER_ID,
    DEFAULT_CLASS_ID,
    TENANT_ID,
    DML_TYPE,
    FIRST_STUDENT_ROW,
    TEMPLATE_COL_A_ROW1,
    TEMPLATE_COL_A_ROW2,
    STUDENT_SCAN_SAFETY_CAP,
    OBJECTIVE_TYPES,
    MCQ_TYPES,
    TF_TYPES,
    DB_CONFIG,
)


# ==================================================
# LOW-LEVEL HELPERS
# ==================================================

def sql_value(value):
    """Convert a Python value into a safe MySQL SQL literal (for .sql files)."""

    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"Invalid numeric value: {value}")
        return str(int(value)) if value.is_integer() else str(value)

    # Strings: escape backslashes and single quotes.
    value = str(value).replace("\\", "\\\\").replace("'", "''")
    return f"'{value}'"


def db_connect():
    """Open a fresh MySQL connection, with a clear message if it fails."""
    try:
        return mysql.connector.connect(**DB_CONFIG)
    except mysql.connector.Error as e:
        raise SystemExit(
            f"Could not connect to MySQL ({DB_CONFIG['host']}:{DB_CONFIG['port']}/"
            f"{DB_CONFIG['database']}): {e}\n"
            f"Check that MySQL is running and the credentials in config.py / env vars are correct."
        )


def derive_section(sheet_name):
    """
    Derive the section letter from a sheet name by taking its leading letters.
        "A1" -> "A",  "B1" -> "B",  "SectionC" -> "SectionC",  "A" -> "A"
    Override is possible via the --sections CLI flag.
    """
    match = re.match(r"^[A-Za-z]+", sheet_name)
    return match.group(0) if match else sheet_name


# ==================================================
# SHEET AUTO-DETECTION
# ==================================================

def detect_template_sheets(workbook):
    """
    Return all worksheet names that follow the template
    (A1 == 'competency' and A2 == 'Students'). Order = workbook order.
    """
    found = []
    for name in workbook.sheetnames:
        ws = workbook[name]
        a1 = ws.cell(1, 1).value
        a2 = ws.cell(2, 1).value
        if (a1 is not None and str(a1).strip().lower() == TEMPLATE_COL_A_ROW1
                and a2 is not None and str(a2).strip().lower() == TEMPLATE_COL_A_ROW2):
            found.append(name)
    return found


def _count_students(ws):
    """Count non-empty student-id cells (col A) from FIRST_STUDENT_ROW until first blank."""
    count = 0
    for row in range(FIRST_STUDENT_ROW, FIRST_STUDENT_ROW + STUDENT_SCAN_SAFETY_CAP):
        v = ws.cell(row, 1).value
        if v is None or v == "":
            break
        count += 1
    return count


def pick_one_sheet_per_section(workbook):
    """
    Detect template sheets and pick ONE per section.

    A workbook may contain several sheets that map to the same section letter
    (e.g. A, A1, A2 all -> section "A"). Only one should be loaded to avoid
    double-counting students. We pick the sheet with the most students
    (first on tie), which naturally selects the "full" sheet (A1/B1/C1/D1)
    over subset/working sheets (A/A2/...).
    """
    detected = detect_template_sheets(workbook)

    # Group detected sheet names by their derived section letter.
    by_section = {}
    for name in detected:
        sec = derive_section(name)
        by_section.setdefault(sec, []).append(name)

    # For each section, pick the sheet with the most students.
    picked = []
    for sec, names in by_section.items():
        if len(names) == 1:
            picked.append(names[0])
            continue
        best_name, best_count = None, -1
        for name in names:
            c = _count_students(workbook[name])
            if c > best_count:          # strict > keeps the first on tie
                best_name, best_count = name, c
        picked.append(best_name)

    # Return in workbook order.
    order = {n: i for i, n in enumerate(detected)}
    picked.sort(key=lambda n: order.get(n, len(detected)))
    return picked


# ==================================================
# STEP 2: FETCH QUESTIONS FROM MYSQL (type + marks dynamic)
# ==================================================

def get_questions(connection, paper_id):
    """
    Return a list of {sequence, qid, quesType, maxMarks, canswer} for every
    question on `paper_id`, ordered by sequence. The number of questions is
    whatever the DB contains - not a fixed constant.
    """

    cursor = connection.cursor(dictionary=True)

    # Join examquestion (sequence + marks + questionId) with questions (type + answer).
    query = """
        SELECT
            eq.questionId AS qid,
            eq.sequence   AS sequence,
            eq.marks      AS maxMarks,
            q.quesType    AS quesType,
            q.correctAns  AS canswer
        FROM examquestion AS eq
        INNER JOIN questions AS q
            ON q.id = eq.questionId
        WHERE eq.paperId = %s
        ORDER BY eq.sequence
    """

    try:
        cursor.execute(query, (paper_id,))
        questions = cursor.fetchall()
    finally:
        cursor.close()

    if not questions:
        raise ValueError(f"Paper {paper_id} has no questions in examquestion.")

    # Index by sequence; detect duplicates / gaps. Sequences need not be 1..N
    # contiguously, but they must be unique and positive.
    by_seq = {}
    for q in questions:
        seq = q["sequence"]
        if seq is None:
            raise ValueError("A question has a NULL sequence.")
        seq = int(seq)
        if seq < 1:
            raise ValueError(f"Unexpected question sequence: {seq}")
        if seq in by_seq:
            raise ValueError(f"Duplicate question sequence: {seq}")
        q["sequence"] = seq
        # maxMarks may be NULL in the DB; treat as unknown (None).
        q["maxMarks"] = None if q["maxMarks"] is None else int(q["maxMarks"])
        q["quesType"] = (q["quesType"] or "").strip()
        by_seq[seq] = q

    # Return sorted by sequence.
    return [by_seq[s] for s in sorted(by_seq)]


# ==================================================
# STEP 1: READ STUDENTS + MARKS FROM EXCEL (dynamic counts)
# ==================================================

def get_students_and_marks(excel_path, sheet_name, expected_sequences):
    """
    Read one worksheet and return a list of {"student_id", "marks": {seq: marks}}.

    `expected_sequences` is the set of question sequences from the DB. The Excel
    must supply a marks column for each of them (cols B, C, ... = sequence 1, 2, ...).
    The student row scan stops at the first blank student-id cell, so the number
    of students is whatever the sheet contains.
    """

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path.resolve()}")

    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)

    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet '{sheet_name}' not found.\n"
                f"Available: {workbook.sheetnames}"
            )

        sheet = workbook[sheet_name]

        # Validate that Excel row 2 has a header for every expected question.
        # Headers are in columns 2..(N+1). We just check they are non-empty.
        max_seq = max(expected_sequences)
        for seq in expected_sequences:
            header = sheet.cell(2, seq + 1).value
            if header is None or str(header).strip() == "":
                raise ValueError(
                    f"Excel worksheet '{sheet_name}' is missing the column "
                    f"for question sequence {seq} (row 2, col {seq + 1})."
                )

        students = []
        seen = set()

        # Scan student rows from FIRST_STUDENT_ROW until a blank id, with a safety cap.
        row = FIRST_STUDENT_ROW
        while row <= FIRST_STUDENT_ROW + STUDENT_SCAN_SAFETY_CAP:
            raw_id = sheet.cell(row, 1).value

            # Stop at the first blank student-id cell.
            if raw_id is None or raw_id == "":
                break

            if isinstance(raw_id, bool):
                raise ValueError(f"Invalid student ID at Excel row {row}.")

            # Excel may store ids as text or number; coerce to int.
            try:
                numeric = float(raw_id)
            except (TypeError, ValueError):
                raise ValueError(f"Invalid student ID {raw_id!r} at Excel row {row}.")

            if not math.isfinite(numeric) or not numeric.is_integer():
                raise ValueError(f"Student ID must be an integer. Row {row}: {raw_id!r}")

            student_id = int(numeric)
            if student_id in seen:
                raise ValueError(f"Duplicate student ID {student_id} in '{sheet_name}'.")
            seen.add(student_id)

            # Read marks for every expected question sequence.
            marks_by_sequence = {}
            for seq in expected_sequences:
                value = sheet.cell(row, seq + 1).value
                if value is None or value == "":
                    value = 0
                if isinstance(value, bool):
                    raise ValueError(
                        f"Invalid marks for student {student_id}, "
                        f"question {seq}: {value!r}"
                    )
                try:
                    marks = float(value)
                except (TypeError, ValueError):
                    raise ValueError(
                        f"Invalid marks for student {student_id}, "
                        f"question {seq}: {value!r}"
                    )
                if not math.isfinite(marks) or marks < 0:
                    raise ValueError(
                        f"Invalid marks for student {student_id}, "
                        f"question {seq}: {value!r}"
                    )
                marks_by_sequence[seq] = int(marks) if marks.is_integer() else marks

            students.append({"student_id": student_id, "marks": marks_by_sequence})
            row += 1

        if not students:
            raise ValueError(f"No students found in worksheet '{sheet_name}'.")

        print(f"Worksheet: {sheet_name}")
        print(f"Student IDs found: {sorted(seen)}")
        return students

    finally:
        workbook.close()


def get_competencies_from_excel(excel_path, sheet_name, sequences):
    """Return {sequence: competency_code} from row 1 of the worksheet (used by --export-meta)."""
    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        sheet = workbook[sheet_name]
        comps = {}
        for seq in sequences:
            v = sheet.cell(1, seq + 1).value
            comps[seq] = None if v is None or str(v).strip() == "" else str(v).strip()
        return comps
    finally:
        workbook.close()


# ==================================================
# STEP 3: BUILD ANSWER ROWS (type-driven, not sequence-driven)
# ==================================================
# STEP 3: BUILD ANSWER ROWS (type-driven, not sequence-driven)
# ==================================================

def build_records(students, questions, section, paper_id, class_id):
    """
    Build one tuple per (student, question) matching the `answer` table columns.
    The answer-fill behaviour is driven by each question's real quesType from the
    DB - so any question pattern works without code changes.
    """

    rows = []

    for student in students:
        student_id = student["student_id"]

        for q in questions:
            seq = q["sequence"]
            qid = q["qid"]
            qtype = q["quesType"]
            correct_answer = q["canswer"]
            marks = student["marks"][seq]

            uanswer, canswer = None, None  # default: open-ended -> NULL

            # Only objective types get answer fields filled.
            if qtype in OBJECTIVE_TYPES:
                if correct_answer is None:
                    raise ValueError(
                        f"Question ID {qid} (type {qtype}) has no correct answer."
                    )
                correct_answer = str(correct_answer).strip()

                if marks > 0:
                    # Scored -> record the correct answer as the student's answer.
                    user_answer = correct_answer
                elif qtype in MCQ_TYPES:
                    # MCQ, zero marks -> record a DIFFERENT option (a wrong one).
                    options = ["optA", "optB", "optC", "optD"]
                    user_answer = next(
                        (o for o in options
                         if o.casefold() != correct_answer.casefold()),
                        "optA",
                    )
                elif qtype in TF_TYPES:
                    # True/False, zero marks -> record the opposite.
                    norm = correct_answer.lower()
                    if norm not in ("true", "false"):
                        raise ValueError(
                            f"Question {qid} (trueFalse) must have "
                            f"'true'/'false' as correct answer. Found: {correct_answer!r}"
                        )
                    user_answer = "true" if norm == "false" else "false"
                else:
                    # Unknown objective type -> fall back to correct answer.
                    user_answer = correct_answer

                uanswer, canswer = user_answer, correct_answer

            # marks stored as string to match the varchar column.
            rows.append((
                paper_id,
                class_id,
                section,
                str(student_id),
                qid,
                uanswer,
                canswer,
                str(marks),
                seq,
                DML_TYPE,
                TENANT_ID,
            ))

    return rows


# ==================================================
# OPTIONAL: RENDER A .sql BACKUP FILE
# ==================================================

def build_sql_text(students, questions, section, sheet_name, rows, paper_id, class_id):
    """Render a full INSERT statement as text (for --save-sql backup files)."""

    record_strings = []
    for r in rows:
        (pid, cid, sec, student_id, qid,
         uanswer, canswer, marks, seq, dml_type, tenant_id) = r
        record_strings.append(
            f"({pid}, {cid}, {sql_value(sec)}, "
            f"{sql_value(student_id)}, {sql_value(qid)}, "
            f"{sql_value(uanswer)}, {sql_value(canswer)}, "
            f"NOW(), {sql_value(marks)}, {seq}, "
            f"{sql_value(dml_type)}, {sql_value(tenant_id)})"
        )

    return (
        "-- Generated from Excel worksheet and MySQL question mapping.\n"
        f"-- Paper ID: {paper_id}; Class: {class_id}; Section: {section}\n"
        f"-- Worksheet: {sheet_name}\n"
        f"-- Students: {len(students)}\n"
        f"-- Questions per student: {len(questions)}\n"
        f"-- Total records: {len(rows)}\n"
        "-- Objective questions: answer fields filled by type-driven logic.\n"
        "-- Open-ended questions: uanswer and canswer are NULL.\n\n"
        "INSERT INTO answer (\n"
        "    paperid, class, section, studentid, qid,\n"
        "    uanswer, canswer, createdate, marks,\n"
        "    sequence, dmlType, tenantid\n"
        ")\nVALUES\n"
        + ",\n".join(record_strings) + ";\n"
    )


# ==================================================
# STEP 4: DATABASE WRITE HELPERS
# ==================================================

def existing_count(connection, paper_id, section):
    """How many answer rows already exist for this paper + section."""
    cursor = connection.cursor()
    try:
        cursor.execute(
            "SELECT COUNT(*) FROM answer WHERE paperid = %s AND section = %s",
            (paper_id, section),
        )
        return cursor.fetchone()[0]
    finally:
        cursor.close()


def delete_existing(connection, paper_id, section):
    """Delete all existing rows for this paper + section (used by --overwrite)."""
    cursor = connection.cursor()
    try:
        cursor.execute(
            "DELETE FROM answer WHERE paperid = %s AND section = %s",
            (paper_id, section),
        )
        deleted = cursor.rowcount
        connection.commit()
        return deleted
    finally:
        cursor.close()


def insert_rows(connection, rows):
    """Insert all rows in one batched transaction. Rolls back on any error."""
    cursor = connection.cursor()
    try:
        cursor.executemany(
            """
            INSERT INTO answer (
                paperid, class, section, studentid, qid,
                uanswer, canswer, createdate, marks,
                sequence, dmlType, tenantid
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, NOW(), %s, %s, %s, %s)
            """,
            rows,
        )
        connection.commit()
        return cursor.rowcount
    except Exception:
        connection.rollback()
        raise
    finally:
        cursor.close()


# ==================================================
# EXPORT QUESTION METADATA (for SQL generation in the dashboard)
# ==================================================
# Writes paper_<id>_meta.json - a small "DB bundle" with the paper/class/tenant
# ids plus each question's qid, type, max marks and correct answer.
#
# WHO USES IT: only someone WITH database access runs `--export-meta` and
# shares the resulting JSON with whoever generates SQL in the dashboard.
# Graph viewers never need this file, so they can never generate SQL.
#
# SECURITY: this file contains NO database credentials - only question
# data (including the correct-answer key). The dashboard uses it to
# generate INSERT *text*; it never connects to MySQL itself.
# ==================================================

def export_meta(excel_path, paper_id, class_id, sheet_for_competency):
    """Write paper_<id>_meta.json - the DB bundle for the dashboard's Generate SQL tab."""
    connection = db_connect()
    try:
        questions = get_questions(connection, paper_id)
    finally:
        connection.close()

    sequences = [q["sequence"] for q in questions]
    comps = get_competencies_from_excel(excel_path, sheet_for_competency, sequences)

    bundle = {
        "paperId": paper_id,
        "classId": class_id,
        "tenantId": TENANT_ID,
        "dmlType": DML_TYPE,
        "questions": [{
            "sequence": q["sequence"],
            "qid": q["qid"],
            "quesType": q["quesType"],
            "maxMarks": q["maxMarks"],
            "correctAns": q["canswer"],
            "competency": comps.get(q["sequence"]),
        } for q in questions],
    }

    out = Path(f"paper_{paper_id}_meta.json")
    out.write_text(json.dumps(bundle, indent=2), encoding="utf-8")
    print(f"Metadata written to {out.resolve()} ({len(bundle['questions'])} questions).")
    print("Share this file only with whoever generates SQL in the dashboard.")
    return out


# ==================================================
# ORCHESTRATE ONE SHEET
# ==================================================

def process_sheet(excel_path, sheet_name, section, *, paper_id, class_id,
                  auto_insert, overwrite, save_sql):
    """Process a single worksheet end-to-end and return a summary dict."""

    # Step 2 (questions first so we know expected sequences & types).
    connection = db_connect()
    try:
        questions = get_questions(connection, paper_id)
        sequences = [q["sequence"] for q in questions]
        existing = existing_count(connection, paper_id, section)

        if existing > 0:
            if overwrite:
                print(f"  Overwrite: deleting {existing} existing rows...")
                deleted = delete_existing(connection, paper_id, section)
                print(f"  Deleted {deleted} rows.")
            elif auto_insert:
                print(
                    f"  WARNING: {existing} rows already exist for section {section}. "
                    f"Skipping insert. Use --overwrite to replace."
                )
                auto_insert = False
            else:
                print(f"  NOTE: {existing} rows already exist (no insert requested).")
    finally:
        connection.close()

    # Step 1: read students + marks (validated against the DB sequences).
    students = get_students_and_marks(excel_path, sheet_name, sequences)

    # Step 3: build rows (type-driven).
    rows = build_records(students, questions, section, paper_id, class_id)

    # Optional .sql backup.
    if save_sql:
        sql_text = build_sql_text(students, questions, section, sheet_name,
                                  rows, paper_id, class_id)
        output_path = Path(f"section_{section}_answers.sql")
        output_path.write_text(sql_text, encoding="utf-8")
        print(f"  SQL file: {output_path.resolve()}")

    # Step 4: insert (transactional).
    if auto_insert:
        connection = db_connect()
        try:
            inserted = insert_rows(connection, rows)
            print(f"  Inserted {inserted} rows into answer table.")
        finally:
            connection.close()

    print(
        f"  Section {section}: {len(students)} students, "
        f"{len(questions)} questions, {len(rows)} records."
    )
    return {
        "section": section,
        "sheet": sheet_name,
        "students": len(students),
        "records": len(rows),
        "existing_before": existing,
    }


# ==================================================
# COMMAND-LINE ENTRY POINT
# ==================================================

def parse_section_override(raw):
    """Parse a --sections value like 'A1=A,B1=B' into {sheet: section}."""
    if not raw:
        return {}
    out = {}
    for pair in raw.split(","):
        if "=" not in pair:
            raise ValueError(f"--sections expects SheetName=Section, got '{pair}'.")
        sheet, sec = pair.split("=", 1)
        out[sheet.strip()] = sec.strip()
    return out


def main():
    """Parse command-line flags and run the requested sheets."""

    parser = argparse.ArgumentParser(
        description=(
            "Insert answer rows from an Excel workbook into MySQL, fully dynamic. "
            "Question count, types, marks, competencies and sheets are all "
            "auto-detected from the workbook and database."
        )
    )

    parser.add_argument("--excel", default="SCIENCE_ROHIT_8 (1) .xlsx",
                        help="Excel workbook path")
    parser.add_argument("--paper", type=int, default=DEFAULT_PAPER_ID,
                        help=f"Paper id (default {DEFAULT_PAPER_ID})")
    parser.add_argument("--class", dest="class_id", type=int, default=DEFAULT_CLASS_ID,
                        help=f"Class id (default {DEFAULT_CLASS_ID})")
    parser.add_argument("--sheet", default=None,
                        help="Process a single worksheet by name (e.g. B1)")
    parser.add_argument("--sheets", default=None,
                        help="Comma-separated list of worksheet names (e.g. A1,B1)")
    parser.add_argument("--sections", default=None,
                        help="Override section per sheet, e.g. 'A1=A,B1=B'")
    parser.add_argument("--all", action="store_true",
                        help="Process all auto-detected template sheets")
    parser.add_argument("--auto-insert", action="store_true", default=True,
                        help="Insert records directly into MySQL (default: on)")
    parser.add_argument("--no-insert", dest="auto_insert", action="store_false",
                        help="Do not insert; only generate SQL files")
    parser.add_argument("--overwrite", action="store_true",
                        help="Delete existing rows for the paper+section before inserting")
    parser.add_argument("--save-sql", action="store_true",
                        help="Also write a .sql backup file per section")
    parser.add_argument("--export-meta", action="store_true",
                        help="Write paper_<id>_meta.json (DB bundle for the dashboard's Generate SQL tab)")

    args = parser.parse_args()

    excel_path = Path(args.excel)
    section_override = parse_section_override(args.sections)

    # Validate the Excel file up front with a clear message (don't crash raw).
    if not excel_path.exists():
        raise SystemExit(f"Excel file not found: {excel_path.resolve()}")

    # --export-meta: write the DB bundle and exit. Only someone with DB
    # access runs this; the resulting JSON is shared with the SQL generator.
    if args.export_meta:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        try:
            sheets = detect_template_sheets(wb)
        finally:
            wb.close()
        if not sheets:
            raise ValueError("No template sheets found in the workbook.")
        export_meta(excel_path, args.paper, args.class_id, sheets[0])
        return

    # Decide which sheets to run.
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        detected = detect_template_sheets(wb)
        default_sheets = pick_one_sheet_per_section(wb)
    finally:
        wb.close()

    if args.sheet:
        sheets = [args.sheet]
    elif args.sheets:
        sheets = [s.strip() for s in args.sheets.split(",")]
    elif args.all or (not args.sheet and not args.sheets):
        sheets = default_sheets  # default: one sheet per section (A1/B1/C1/D1)
    else:
        sheets = default_sheets

    # Validate chosen sheets exist and are template sheets.
    unknown = [s for s in sheets if s not in detected]
    if unknown:
        print(f"WARNING: these sheets were not detected as template sheets: {unknown}")
        print(f"         Detected template sheets: {detected}")

    # Resolve a section for each sheet (override > derive from name).
    resolved = []
    seen_sections = set()
    for s in sheets:
        sec = section_override.get(s, derive_section(s))
        resolved.append((s, sec))
        if sec in seen_sections:
            print(f"WARNING: multiple sheets map to section '{sec}' "
                  f"(sheet {s}). Inserts will share that section.")
        seen_sections.add(sec)

    print(f"Paper ID: {args.paper} | Class: {args.class_id}")
    print(f"Sheets: {[s for s,_ in resolved]}  ->  sections: {[c for _,c in resolved]}")
    print(f"Auto-insert: {args.auto_insert} | Overwrite: {args.overwrite} | "
          f"Save SQL: {args.save_sql}")
    print("-" * 50)

    # Run each sheet. One bad sheet prints an error and the rest still run.
    results = []
    for sheet_name, section in resolved:
        print(f"\nProcessing {sheet_name} (section {section}) ...")
        try:
            results.append(process_sheet(
                excel_path, sheet_name, section,
                paper_id=args.paper,
                class_id=args.class_id,
                auto_insert=args.auto_insert,
                overwrite=args.overwrite,
                save_sql=args.save_sql,
            ))
        except SystemExit:
            raise  # fatal errors (DB connection, missing file) still stop the run
        except Exception as e:
            # A data problem in one sheet should not kill the whole run.
            print(f"  ERROR processing {sheet_name}: {e}")
            print(f"  Skipped {sheet_name}. Continuing with the remaining sheets.")
            results.append({
                "section": section, "sheet": sheet_name,
                "students": 0, "records": 0, "existing_before": 0, "error": str(e),
            })

    # Final summary.
    print("\n" + "=" * 50)
    print("SUMMARY")
    print("=" * 50)
    for r in results:
        print(
            f"  Section {r['section']} ({r['sheet']}): "
            f"{r['students']} students, {r['records']} records, "
            f"{r['existing_before']} pre-existing"
        )
    print("Done.")


# Run main() only when executed directly, not when imported.
if __name__ == "__main__":
    main()