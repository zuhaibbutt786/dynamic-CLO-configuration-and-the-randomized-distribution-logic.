import streamlit as st
import pandas as pd
import random
import io
import re
import openpyxl
from openpyxl.utils import get_column_letter

# --- Page Configuration ---
st.set_page_config(
    page_title="OBE Intelligence Pro",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom Styling ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; font-weight: bold; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; background-color: #28a745; color: white; font-weight: bold; }
    div[data-testid="stExpander"] { border: 1px solid #e1e4e8; border-radius: 8px; background-color: white; }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; }
    .stTabs [data-baseweb="tab-list"] { gap: 16px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #f0f2f6; border-radius: 5px 5px 0 0; padding: 10px 16px; }
    .stTabs [aria-selected="true"] { background-color: #007bff !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)


# --- Core Logic ---
def distribute_marks(obtained, clo_max_marks):
    """Robust mark distribution logic with error handling."""
    try:
        obtained = float(obtained)
        if pd.isna(obtained) or obtained <= 0:
            return [0.0] * len(clo_max_marks)

        total_possible = sum(clo_max_marks)
        if obtained > total_possible:
            obtained = total_possible

        allocated = [0.0] * len(clo_max_marks)
        remaining = obtained

        iterations = 0
        while remaining > 0.001 and iterations < 5000:
            available_indices = [i for i, max_val in enumerate(clo_max_marks) if allocated[i] < max_val]
            if not available_indices:
                break

            idx = random.choice(available_indices)
            space_left = clo_max_marks[idx] - allocated[idx]

            step = random.choice([0.5, 1.0])
            chunk = min(step, remaining, space_left)

            allocated[idx] += chunk
            remaining -= chunk
            iterations += 1

        return [round(m, 2) for m in allocated]
    except Exception:
        return [0.0] * len(clo_max_marks)


def parse_max_from_header(header_text):
    """Extract numeric max from strings like 'Q1 (10.00)' or 'Q5 (10.00)\\n'."""
    if header_text is None:
        return 10.0
    m = re.search(r"\(([\d.]+)\)", str(header_text))
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return 10.0


def safe_numeric(val):
    """Convert cell value to float; treat 'A'/absent/blank as None."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().upper()
    if s in ("", "A", "ABSENT", "N/A", "-", "NONE"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def load_uis_marks(uploaded_file):
    """
    Load UIS-style marks workbook.
    Expects header row with Q1/Q2/.../A1/.../Mid1/Final1/CP1...
    and a RollNo / Name column.
    Returns a DataFrame indexed by normalized RollNo string.
    """
    # Try reading with openpyxl first to find the header row
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    header_row_idx = None
    for r in range(1, min(20, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, min(20, ws.max_column + 1))]
        if any(v in ("ROLLNO", "ROLL NO", "ROLL_NO", "REGISTRATION NO.", "REG NO") for v in row_vals) or "SR#" in row_vals:
            # Prefer the row that also has numeric maxes or short codes nearby
            header_row_idx = r
            # Often the short codes (Q1, A1, Mid1) are one row above the max marks row
            break

    uploaded_file.seek(0)

    # Fallback: read all and let user map, but try smart defaults
    df_raw = pd.read_excel(uploaded_file, header=None, engine="openpyxl")

    # Find best header row: one containing RollNo-like and short component codes
    best_r = None
    best_score = -1
    for r in range(min(15, len(df_raw))):
        vals = [str(x).strip().upper() if pd.notna(x) else "" for x in df_raw.iloc[r].tolist()]
        score = 0
        if any("ROLL" in v or v == "SR#" or "REG" in v for v in vals):
            score += 3
        if any(v in ("Q1", "Q2", "A1", "MID1", "FINAL1", "CP1") for v in vals):
            score += 5
        if any(re.fullmatch(r"\d+(\.\d+)?", v) for v in vals):
            score += 1  # max-marks row
        if score > best_score:
            best_score = score
            best_r = r

    if best_r is None:
        best_r = 0

    # Component codes may be on best_r, max marks on best_r+1, or vice versa
    codes_row = best_r
    max_row = best_r + 1 if best_r + 1 < len(df_raw) else best_r

    # Prefer the row that has more short alpha codes as the code row
    def code_score(row_idx):
        vals = [str(x).strip().upper() if pd.notna(x) else "" for x in df_raw.iloc[row_idx].tolist()]
        return sum(1 for v in vals if re.fullmatch(r"[A-Z]+\d*", v) and len(v) <= 8)

    if code_score(max_row) > code_score(codes_row):
        codes_row, max_row = max_row, codes_row

    headers = []
    for c in range(df_raw.shape[1]):
        code = str(df_raw.iloc[codes_row, c]).strip() if pd.notna(df_raw.iloc[codes_row, c]) else ""
        mx = str(df_raw.iloc[max_row, c]).strip() if pd.notna(df_raw.iloc[max_row, c]) else ""
        if code and re.fullmatch(r"[A-Za-z]+\d*", code):
            headers.append(code.upper())
        elif code.upper() in ("SR#", "ROLLNO", "ROLL NO", "NAME", "REGISTRATION NO."):
            headers.append(code.upper().replace(" ", ""))
        elif mx and re.fullmatch(r"\d+(\.\d+)?", mx) and not code:
            headers.append(f"COL{c}")
        else:
            headers.append(code if code else f"COL{c}")

    # Data starts after the max-marks row
    data_start = max(codes_row, max_row) + 1
    df = df_raw.iloc[data_start:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    # Normalize roll column name
    roll_candidates = [c for c in df.columns if "ROLL" in str(c).upper() or str(c).upper() in ("REGNO", "REGISTRATIONNO.", "REGISTRATIONNO")]
    if not roll_candidates:
        # sometimes first or second col
        roll_candidates = [df.columns[1]] if len(df.columns) > 1 else [df.columns[0]]
    roll_col = roll_candidates[0]
    df["_ROLL"] = df[roll_col].astype(str).str.strip()

    # Drop average / total footer rows
    df = df[~df["_ROLL"].str.upper().isin(("AVERAGE", "AVG", "TOTAL", "NAN", "NONE"))]
    df = df[df["_ROLL"].str.len() > 3]

    return df.set_index("_ROLL", drop=False)


def detect_qobe_structure(ws):
    """
    Parse QOBE template header rows.
    Returns list of dicts: {col, activity, question, max_marks}
    and student start row, roll col, name col.
    """
    # Find header row with 'Registration No.' or 'Q1'
    header_row = 4
    for r in range(1, 10):
        v = str(ws.cell(r, 1).value or "").strip().lower()
        if "registration" in v or "roll" in v:
            header_row = r
            break

    activity_row = header_row - 1
    columns_meta = []
    for c in range(3, ws.max_column + 1):
        activity = str(ws.cell(activity_row, c).value or "").strip()
        q_header = str(ws.cell(header_row, c).value or "").strip()
        mx = parse_max_from_header(q_header)
        q_label = q_header.split("(")[0].strip() if q_header else f"Q{c}"
        columns_meta.append({
            "col": c,
            "letter": get_column_letter(c),
            "activity": activity,
            "question": q_label,
            "max_marks": mx,
            "header": q_header,
        })

    # Group consecutive columns that share the same activity name
    groups = []
    current = None
    for meta in columns_meta:
        act = meta["activity"] or f"COL{meta['col']}"
        if current is None or current["activity"] != act:
            current = {"activity": act, "cols": [meta]}
            groups.append(current)
        else:
            current["cols"].append(meta)

    student_start = header_row + 1
    return {
        "header_row": header_row,
        "student_start": student_start,
        "columns_meta": columns_meta,
        "groups": groups,
        "roll_col": 1,
        "name_col": 2,
    }


def main():
    st.title("🎯 OBE Marks Intelligence Pro")
    st.markdown("##### Transform portal exports into structured CLO data and map them to official QOBE templates.")

    with st.sidebar:
        st.header("Help & Instructions")
        st.info("""
        **Tab 1 – Distribution Engine**  
        Split one total mark across CLOs.

        **Tab 2 – Template Mapper**  
        Write distributed CLO columns into a generic template.

        **Tab 3 – Bulk QOBE Update** ⭐  
        Upload UIS marks file + empty QOBE Activity Outcome template.  
        Single-question activities are copied; Mid/Final totals are randomly split across their Q columns.
        """)
        st.divider()
        st.caption("v2.2 – Bulk QOBE Update")

    tab1, tab2, tab3 = st.tabs([
        "📊 1. Distribution Engine",
        "📄 2. Template Mapper",
        "⚡ 3. Bulk QOBE Update",
    ])

    # ========== TAB 1: DISTRIBUTION ==========
    with tab1:
        st.header("Step 1: Data Intake")
        uploaded_file = st.file_uploader(
            "Upload Roster / Result File",
            type=["csv", "xlsx", "xls"],
            help="Upload the file from your portal. Even if it says .xls, this tool will fix format mismatches.",
            key="tab1_upload",
        )

        if uploaded_file:
            try:
                file_ext = uploaded_file.name.split(".")[-1].lower()

                if file_ext == "csv":
                    df = pd.read_csv(uploaded_file)
                else:
                    try:
                        engine = "xlrd" if file_ext == "xls" else "openpyxl"
                        df = pd.read_excel(uploaded_file, engine=engine)
                    except Exception as e:
                        error_msg = str(e).lower()
                        if any(x in error_msg for x in [
                            "bof", "html", "unsupported format", "xlrd",
                            "openpyxl", "no text parsed", "expected bof",
                            "file is not a zip", "workbook",
                        ]):
                            st.info("🔄 Technical format mismatch detected. Activating HTML Recovery Mode...")
                            uploaded_file.seek(0)
                            content = uploaded_file.read()
                            try:
                                html_content = content.decode("utf-8")
                            except UnicodeDecodeError:
                                html_content = content.decode("latin-1")
                            html_tables = pd.read_html(io.StringIO(html_content))
                            if not html_tables:
                                raise ValueError("No tables found in the HTML content.")
                            df = max(html_tables, key=lambda t: t.shape[0] * t.shape[1])
                            df.columns = [str(c).strip() for c in df.columns]
                        else:
                            raise e

                df = df.dropna(how="all").dropna(axis=1, how="all")
                st.success(f"✅ Successfully loaded {len(df)} records.")

                with st.expander("🔍 Preview Raw Portal Data"):
                    st.dataframe(df.head(10), use_container_width=True)

                st.divider()
                st.header("Step 2: Column Mapping & CLO Setup")
                cols = df.columns.tolist()

                c1, c2, c3 = st.columns(3)
                name_col = c1.selectbox("Name Column", cols, help="Column containing student names.")
                roll_col = c2.selectbox("Roll No Column", cols, help="Column containing Registration/IDs.")
                marks_col = c3.selectbox("Obtained Marks", cols, help="The total score column you want to split.")

                num_clos = st.number_input("Total CLOs in this Exam", min_value=1, value=3, help="Define how many CLOs were tested.")

                clo_max_marks = []
                clo_ui_cols = st.columns(num_clos)
                for i in range(num_clos):
                    m = clo_ui_cols[i].number_input(
                        f"CLO {i+1} Max", min_value=0.1, value=10.0, key=f"m_{i}",
                        help=f"Set max possible marks for CLO {i+1}",
                    )
                    clo_max_marks.append(m)

                if st.button("🚀 Generate CLO Distribution"):
                    with st.spinner("Calculating distributions..."):
                        if not pd.api.types.is_numeric_dtype(df[marks_col]):
                            df[marks_col] = pd.to_numeric(df[marks_col], errors="coerce").fillna(0)

                        results = df[marks_col].apply(lambda x: distribute_marks(x, clo_max_marks))
                        for i in range(num_clos):
                            df[f"CLO_{i+1}_GEN"] = [res[i] for res in results]

                        st.session_state["processed_df"] = df
                        st.session_state["num_clos"] = num_clos
                        st.session_state["name_col"] = name_col
                        st.session_state["roll_col"] = roll_col
                        st.balloons()
                        st.success("Marks Distributed! Proceed to 'Template Mapper' tab.")
                        st.dataframe(df.head(), use_container_width=True)

            except Exception as e:
                st.error(f"❌ Error reading file: {e}")

    # ========== TAB 2: TEMPLATE MAPPER ==========
    with tab2:
        if "processed_df" not in st.session_state:
            st.warning("⚠️ Please process your data in the 'Distribution Engine' tab first.")
        else:
            st.header("Step 3: Official Template Integration")
            template_file = st.file_uploader(
                "Upload Your University Excel Template",
                type=["xlsx"],
                help="Upload the blank Excel sheet where you need the marks filled. MUST be .xlsx.",
                key="tab2_template",
            )

            if template_file:
                try:
                    wb_temp = openpyxl.load_workbook(template_file)
                    sheet_name = st.selectbox("Select Target Sheet", wb_temp.sheetnames)
                    sheet = wb_temp[sheet_name]

                    st.info("📍 Map Excel Coordinates")
                    m1, m2, m3 = st.columns(3)
                    start_row = m1.number_input("Starting Row", min_value=1, value=5, help="Row number where the first student's name appears.")
                    name_target = m2.text_input("Name Col (Letter)", "B").upper()
                    roll_target = m3.text_input("Roll Col (Letter)", "A").upper()

                    st.write("**Map CLOs to Template Columns**")
                    clo_map_cols = st.columns(st.session_state["num_clos"])
                    clo_target_letters = []
                    for i in range(st.session_state["num_clos"]):
                        let = clo_map_cols[i].text_input(
                            f"CLO {i+1} Col", value=get_column_letter(3 + i), key=f"t_{i}"
                        ).upper()
                        clo_target_letters.append(let)

                    if st.button("🪄 Finalize & Map Template"):
                        with st.spinner("Writing to template..."):
                            final_df = st.session_state["processed_df"]
                            name_col = st.session_state.get("name_col")
                            roll_col = st.session_state.get("roll_col")
                            for idx, row in final_df.iterrows():
                                curr = int(start_row + idx)
                                sheet[f"{name_target}{curr}"] = row[name_col]
                                sheet[f"{roll_target}{curr}"] = row[roll_col]
                                for i in range(st.session_state["num_clos"]):
                                    sheet[f"{clo_target_letters[i]}{curr}"] = row[f"CLO_{i+1}_GEN"]

                            out_ptr = io.BytesIO()
                            wb_temp.save(out_ptr)
                            out_ptr.seek(0)

                            st.download_button(
                                label="💾 Download Mapped OBE Result",
                                data=out_ptr,
                                file_name="OBE_Final_Mapping.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                except Exception as e:
                    st.error(f"❌ Template Error: {e}")

    # ========== TAB 3: BULK QOBE UPDATE ==========
    with tab3:
        st.header("⚡ Bulk QOBE Update")
        st.markdown(
            "Upload your **UIS marks file** (with Q1/A1/Mid1/Final1/CP…) and the empty **QOBE Activity Outcome** template. "
            "The tool matches students by Registration/Roll No, copies single-question totals, "
            "and **randomly distributes** Mid & Final totals across their question columns."
        )

        col_u, col_t = st.columns(2)
        with col_u:
            uis_file = st.file_uploader(
                "1️⃣ UIS Marks File (Book123 / portal export)",
                type=["xlsx", "xls", "csv"],
                key="bulk_uis",
            )
        with col_t:
            qobe_file = st.file_uploader(
                "2️⃣ Empty QOBE Activity Outcome Template",
                type=["xlsx"],
                key="bulk_qobe",
            )

        if uis_file and qobe_file:
            try:
                # --- Load UIS ---
                uis_file.seek(0)
                uis_df = load_uis_marks(uis_file)
                st.success(f"✅ UIS loaded: **{len(uis_df)}** students")

                with st.expander("Preview UIS data (first 8 rows)"):
                    st.dataframe(uis_df.head(8), use_container_width=True)

                # --- Load QOBE structure ---
                qobe_file.seek(0)
                wb = openpyxl.load_workbook(qobe_file)
                sheet_name = st.selectbox("QOBE Sheet", wb.sheetnames, key="bulk_sheet")
                ws = wb[sheet_name]
                structure = detect_qobe_structure(ws)

                st.info(
                    f"Detected **{len(structure['groups'])}** activity groups, "
                    f"students start at row **{structure['student_start']}**"
                )

                # --- Mapping UI ---
                st.subheader("Map UIS columns → QOBE activities")
                uis_cols = [c for c in uis_df.columns if c not in ("_ROLL",) and not str(c).startswith("COL")]
                # Prefer short codes
                preferred = [c for c in uis_cols if re.fullmatch(r"[A-Z]+\d*", str(c).upper())]
                other = [c for c in uis_cols if c not in preferred]
                uis_options = ["— skip —"] + preferred + other

                # Smart default guesses
                def guess_uis(activity_name):
                    a = activity_name.upper()
                    guesses = {
                        "ASSIGNMENT 1": "A1",
                        "ASSIGNMENT 2": "A2",
                        "ASSIGNMENT 3": "A3",
                        "QUIZ 1": "Q1",
                        "QUIZ 2": "Q2",
                        "QUIZ 3": "Q3",
                        "MID TERM / SESSIONAL EXAM  1": "MID1",
                        "MID TERM / SESSIONAL EXAM 1": "MID1",
                        "FINAL EXAM 1": "FINAL1",
                        "PROJECT WORK / LAB WORK / CLASS WORK 1": "CP1",
                        "PROJECT WORK / LAB WORK / CLASS WORK 2": "CP2",
                        "PROJECT WORK / LAB WORK / CLASS WORK 3": "CP3",
                    }
                    # normalize spaces
                    for k, v in guesses.items():
                        if k.replace(" ", "") == a.replace(" ", ""):
                            return v
                    # fuzzy
                    if "ASSIGNMENT" in a and "1" in a:
                        return "A1"
                    if "ASSIGNMENT" in a and "2" in a:
                        return "A2"
                    if "ASSIGNMENT" in a and "3" in a:
                        return "A3"
                    if "QUIZ" in a and "1" in a:
                        return "Q1"
                    if "QUIZ" in a and "2" in a:
                        return "Q2"
                    if "QUIZ" in a and "3" in a:
                        return "Q3"
                    if "MID" in a:
                        return "MID1"
                    if "FINAL" in a:
                        return "FINAL1"
                    if "PROJECT" in a or "CLASS WORK" in a or "LAB" in a:
                        if "1" in a:
                            return "CP1"
                        if "2" in a:
                            return "CP2"
                        if "3" in a:
                            return "CP3"
                    return "— skip —"

                mappings = []
                for gi, group in enumerate(structure["groups"]):
                    act = group["activity"]
                    n_q = len(group["cols"])
                    maxes = [c["max_marks"] for c in group["cols"]]
                    default = guess_uis(act)
                    # ensure default is in options
                    if default not in uis_options:
                        default = "— skip —"

                    c1, c2, c3 = st.columns([3, 2, 2])
                    with c1:
                        src = st.selectbox(
                            f"**{act}**  ({n_q} Q → max {maxes})",
                            uis_options,
                            index=uis_options.index(default),
                            key=f"map_{gi}",
                        )
                    with c2:
                        mode = st.radio(
                            "Mode",
                            ["Direct (1 col)", "Distribute"] if n_q > 1 else ["Direct (1 col)"],
                            index=1 if n_q > 1 else 0,
                            key=f"mode_{gi}",
                            horizontal=True,
                        )
                    with c3:
                        st.caption(f"Cols: {', '.join(c['letter'] for c in group['cols'])}")

                    mappings.append({
                        "group": group,
                        "uis_col": src if src != "— skip —" else None,
                        "mode": mode,
                        "maxes": maxes,
                    })

                absent_policy = st.radio(
                    "How to handle Absent ('A') / blank in UIS?",
                    ["Leave QOBE cell empty", "Write 0"],
                    horizontal=True,
                )

                if st.button("🚀 Run Bulk Update & Download", type="primary"):
                    with st.spinner("Matching students and writing marks..."):
                        matched = 0
                        skipped = 0
                        # Build roll → row index in template
                        roll_to_row = {}
                        for r in range(structure["student_start"], ws.max_row + 1):
                            roll_val = ws.cell(r, structure["roll_col"]).value
                            if roll_val is None:
                                continue
                            roll_to_row[str(roll_val).strip()] = r

                        for roll, uis_row in uis_df.iterrows():
                            roll_key = str(roll).strip()
                            if roll_key not in roll_to_row:
                                # try without leading zeros / int form
                                alt = roll_key.lstrip("0") or "0"
                                if alt not in roll_to_row:
                                    skipped += 1
                                    continue
                                roll_key = alt

                            target_row = roll_to_row[roll_key]
                            matched += 1

                            for m in mappings:
                                if not m["uis_col"]:
                                    continue
                                raw = uis_row.get(m["uis_col"])
                                num = safe_numeric(raw)

                                cols = m["group"]["cols"]
                                if num is None:
                                    if absent_policy == "Write 0":
                                        for meta in cols:
                                            ws.cell(target_row, meta["col"]).value = 0
                                    # else leave empty
                                    continue

                                if m["mode"].startswith("Direct") or len(cols) == 1:
                                    # Put full value into first question column only
                                    # (or clamp to that column's max)
                                    val = min(num, cols[0]["max_marks"])
                                    ws.cell(target_row, cols[0]["col"]).value = val
                                    # clear other Qs of this group if any
                                    for meta in cols[1:]:
                                        ws.cell(target_row, meta["col"]).value = None
                                else:
                                    # Distribute across all Q columns of this activity
                                    parts = distribute_marks(num, m["maxes"])
                                    for meta, part in zip(cols, parts):
                                        ws.cell(target_row, meta["col"]).value = part

                        out = io.BytesIO()
                        wb.save(out)
                        out.seek(0)

                        st.success(f"✅ Matched **{matched}** students  |  Unmatched in template: **{skipped}**")
                        st.download_button(
                            label="💾 Download Filled QOBE Template",
                            data=out,
                            file_name="QOBE_Activity_Outcome_FILLED.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

            except Exception as e:
                st.error(f"❌ Bulk update error: {e}")
                st.exception(e)


if __name__ == "__main__":
    main()
