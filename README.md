# dynamic-CLO-configuration-and-the-randomized-distribution-logic.

Randomly distributing a fixed total (the student's obtained marks) across multiple buckets (the CLOs) without exceeding the maximum capacity of any single bucket.

Ready-to-run Streamlit application that handles file upload, column mapping, dynamic CLO configuration, and randomized distribution logic.

**Live app:** [dynamic-clo-configuration-and-the-randomized-distribution-logi.streamlit.app](https://dynamic-clo-configuration-and-the-randomized-distribution-logi.streamlit.app/)

## Features

- Upload `.csv`, `.xlsx`, or `.xls` (including HTML-exported "Excel" files from university portals)
- Automatic HTML recovery mode when the file is not a real binary Excel format
- Dynamic CLO count and max-marks configuration
- Randomized mark distribution that never exceeds any CLO capacity
- Map results into an official university Excel template

## Recent fix (v2.1)

Fixed the error:

```
Technical format mismatch detected. Activating HTML Recovery Mode...
Error reading file: no text parsed from document (line 0)
```

**Cause:** University portals often export HTML tables with a `.xls` extension. After the failed Excel parse, the file pointer was at EOF and `pd.read_html` received an empty stream.

**Solution:**
1. `uploaded_file.seek(0)` + read raw bytes
2. Decode as UTF-8 / Latin-1
3. Parse with `pd.read_html(io.StringIO(...))`
4. Automatically select the largest table (the marks roster)
5. Clean column names

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```
