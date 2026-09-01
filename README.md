# dynamic-CLO-configuration-and-the-randomized-distribution-logic.

Randomly distributing a fixed total (the student's obtained marks) across multiple buckets (the CLOs) without exceeding the maximum capacity of any single bucket.

Ready-to-run Streamlit application for OBE / QOBE mark workflows.

**Live app:** [dynamic-clo-configuration-and-the-randomized-distribution-logi.streamlit.app](https://dynamic-clo-configuration-and-the-randomized-distribution-logi.streamlit.app/)

## Features

- Upload `.csv`, `.xlsx`, or `.xls` (including HTML-exported "Excel" files from university portals)
- Automatic HTML recovery mode when the file is not a real binary Excel format
- Dynamic CLO count and max-marks configuration
- Randomized mark distribution that never exceeds any CLO capacity
- Map results into an official university Excel template
- **Bulk QOBE Update (v2.2)** – fill empty QOBE Activity Outcome templates from UIS marks files

## Tab 3 – Bulk QOBE Update

1. Upload **UIS marks file** (e.g. Book123.xlsx with columns Q1, Q2, A1… Mid1, Final1, CP1…)
2. Upload **empty QOBE Activity Outcome** template (Registration No. + Name + question columns)
3. Confirm auto-detected mapping (Assignment 1 → A1, Mid → Mid1, etc.)
4. Choose mode per activity:
   - **Direct** – copy total into the first Q column (for single-question activities)
   - **Distribute** – randomly split the total across all Q columns of that activity (Mid / Final)
5. Handle Absent (`A`) as empty cell or `0`
6. Download the filled QOBE template, matched by Registration / Roll No.

### Typical mapping (Professional Practices example)

| QOBE Activity | UIS column | Mode |
|---|---|---|
| Assignment 1 / 2 / 3 | A1 / A2 / A3 | Direct |
| Quiz 1 / 2 / 3 | Q1 / Q2 / Q3 | Direct |
| Mid Term (Q1–Q5) | Mid1 | Distribute |
| Final Exam (Q1–Q4) | Final1 | Distribute |
| Class / Project Work 1–3 | CP1 / CP2 / CP3 | Direct |

## Recent changes

### v2.2 – Bulk QOBE Update
New tab that maps a full UIS marks export into an empty QOBE Activity Outcome workbook in one click.

### v2.1 – HTML .xls recovery
Fixed `no text parsed from document (line 0)` when portals export HTML tables as `.xls`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```
