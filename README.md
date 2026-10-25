# Exam Marks Automation & Analytics

Turn one marks Excel into auto-loaded answer rows **and** ~25 live competency graphs — no more manual data entry, and sharper insight than spreadsheet eyeballing.

Built by **Rohit Kumar**.

---

## What's inside

| File | What it does |
|------|-------------|
| `dashboard.html` | The graphs dashboard. Open in any browser, pick the Excel, ~25 graphs appear. No install, no server, nothing uploaded. |
| `script.py` | Excel → MySQL auto-insert. One command loads every student's answer rows directly into the `answer` table. |
| `config.py` | Shared settings (paper id, class id, DB credentials, type→behaviour map). The one file to edit per exam. |
| `requirements.txt` | Python deps (`openpyxl`, `mysql-connector-python`). |
| `Exam_Marks_Automation_and_Analytics_with_Screenshots.docx` | Formal write-up with screenshots. |
| `Team_Message.md` / `Message_for_Leadership.md` | Ready-to-send messages for the team / leadership. |

> The marks workbook (`SCIENCE_ROHIT_8 (1) .xlsx`) and the answer-key bundle (`paper_76_meta.json`) are **not** committed — they contain real student data and correct answers. Provide your own locally.

---

## Quick start

### 1. View the graphs (anyone, no install)
```powershell
start dashboard.html        # then click "Choose Excel" and pick your marks workbook
```

### 2. Insert marks into the database (one command)
```powershell
pip install -r requirements.txt
python script.py --all --overwrite
```

### 3. (Optional) Generate SQL in the browser
```powershell
python script.py --export-meta      # writes paper_<id>_meta.json (once)
```
Then in the dashboard: load the Excel + load `paper_<id>_meta.json` → the **Generate SQL** tab unlocks → pick sections → Generate → Copy / Download `.sql`.

---

## The Excel template (auto-detected)

Every sheet that follows this template is picked up automatically:

- **Row 1, col A** = `competency`; cols B onward = one competency code per question (`R`, `U`, `Ap`, `An`, `Ev`…)
- **Row 2, col A** = `Students`; cols B onward = `Q1`, `Q2`…
- **Rows 3+** = one student per row (col A = student id, cols B onward = marks)
- **Side table (cols 35–37)** = per question: Q number, max marks, type (`MCQ`/`TF`/`WR`)
- **Band table (cols 38–41)** = Need Improvement / Adequate / Sufficient / Excellent with student counts

Sheet names can be anything; the section letter comes from the leading letters (`A1`→`A`, `B1`→`B`…). If multiple sheets map to the same section, the one with the most students is used.

Nothing is hardcoded — question count, competency codes, marks, types, students and bands are all detected from the workbook at run time.

---

## Tabs in the dashboard

1. **Overview** — headline KPIs + score distribution + competency coverage
2. **Competency** — avg % per skill + zero/full-scorers + "which students" table
3. **Questions** — avg % per question + attempt rate + Hardest 5 / Easiest 5
4. **Students** — ranking + Top 10 / Bottom 10 tables
5. **Sections** — section comparison + pass rate
6. **Bands** — performance band distribution + per-section split + Excel-verification
7. **Generate SQL** — (gated) produce INSERT text in the browser

Each graph has a plain-language caption with a concrete example, so a non-technical reader understands it without prior context.

---

## Security

- The dashboard **never connects to the database** — it only reads the Excel in the browser, so no DB credentials are ever exposed.
- The insertion command runs only on the DB machine.
- No SQL-injection surface (the browser only generates inert text, it does not execute anything).
- The "Generate SQL" tab is gated by `paper_<id>_meta.json`, which only a DB-access user can produce — graph viewers never get it, so they can never generate SQL.

---

## Live demo

This dashboard is hosted on GitHub Pages:

**https://rohiterror505.github.io/analytics_Align/**

Open the link, click "Choose Excel", and pick your marks workbook — the graphs render instantly, all in your browser.

---

## License

Internal use.