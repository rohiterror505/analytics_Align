"""
config.py
=========
All shared settings for the project in ONE place.

This file holds ONLY:
  - The Excel template constants (the layout that does NOT change).
  - Default exam identifiers (overridable on the command line).
  - Database connection settings.

Nothing about question count, question types, marks, competency codes, or
sheet names is hardcoded here - those are all detected dynamically from the
workbook and the database at run time. So if a future exam has 40 questions,
different competencies, or a different MCQ/True-False/CR split, nothing here
needs to change.

This file is imported by script.py; it is not run directly.
"""

import os

# ==================================================
# EXCEL TEMPLATE LAYOUT
# The structure of every worksheet (this stays constant across exams):
#   Row 1, col A = the word "competency"; cols B.. = one competency code per question
#   Row 2, col A = the word "Students";   cols B.. = Q1, Q2, ... headers
#   Row 3 onward  = one student per row (col A = student id, cols B.. = marks)
# ==================================================

FIRST_STUDENT_ROW = 3                 # first Excel row that holds a student
TEMPLATE_COL_A_ROW1 = "competency"    # marker used to auto-detect template sheets
TEMPLATE_COL_A_ROW2 = "students"      # marker used to auto-detect template sheets
STUDENT_SCAN_SAFETY_CAP = 500         # never scan more rows than this (sanity bound)

# ==================================================
# DEFAULT EXAM IDENTIFIERS
# Used when you do NOT pass --paper / --class on the command line.
# Override at run time:  python script.py --paper 90 --class 9
# ==================================================

DEFAULT_PAPER_ID = 76  # your exam paper id
DEFAULT_CLASS_ID = 8   # your class id
TENANT_ID = "YOUR_TENANT_ID"  # multi-tenant id written to answer.tenantid
DML_TYPE = "I"                        # "I" = Insert marker used by the software

# ==================================================
# QUESTION-TYPE -> ANSWER-FILL BEHAVIOUR
# The script reads each question's real type from questions.quesType and uses
# this map to decide how to fill answer.uanswer / answer.canswer.
#   "objective"  -> fill uanswer (correct if marks>0, else a wrong one) + canswer
#   "open"       -> leave uanswer and canswer NULL (constructed response etc.)
# Add new types here as your software grows; no other code needs to change.
# ==================================================

# Types whose answer is an objective choice (MCQ / True-False).
OBJECTIVE_TYPES = {"multipleChoice", "trueFalse"}

# Among objective types, which option style do they use?
MCQ_TYPES = {"multipleChoice"}        # answers look like optA / optB / optC / optD
TF_TYPES = {"trueFalse"}              # answers look like true / false
# Every other type (CR, fillups, oneWord, match, ...) is treated as "open".

# ==================================================
# DATABASE CONFIGURATION
# Values can be overridden with environment variables without editing this file:
#     $env:DB_PASSWORD = "secret"
# ==================================================

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": int(os.getenv("DB_PORT", "3306")),
    "user": os.getenv("DB_USER", "admin"),
    "password": os.getenv("DB_PASSWORD", "YOUR_MYSQL_PASSWORD"),
    "database": os.getenv("DB_NAME", "YOUR_MYSQL_DATABASE"),
}
