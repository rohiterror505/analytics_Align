# Exam Marks Analytics Dashboard

A browser-based competency analytics dashboard for exam marks. Open the link, pick your Excel workbook, and ~25 graphs appear instantly — no install, no server, nothing uploaded.

Built by **Rohit Kumar**.

## Live

**https://rohiterror505.github.io/analytics_Align/**

## Use

### View the graphs (anyone)
Open the link above → click **Choose Excel** → pick your marks workbook.

### Insert marks into the database (DB-access user)
```powershell
pip install -r requirements.txt
python script.py --all --overwrite
```

## Files
- `index.html` — the dashboard (the hosted page)
- `script.py` — Excel → MySQL auto-insert
- `config.py` — shared settings
- `requirements.txt` — Python deps

The marks workbook and answer-key JSON are **not** committed (they contain real student data and correct answers). Provide your own locally.