import streamlit as st
import pandas as pd
import random
import io
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
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; background-color: #28a745; color: white; }
    .css-1kyx7g3 { background-color: #ffffff; padding: 2rem; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

# --- Core Logic ---
def distribute_marks(obtained, clo_max_marks):
    """Robust mark distribution logic with error handling."""
    try:
        obtained = float(obtained)
        if pd.isna(obtained) or obtained <= 0:
            return [0.0] * len(clo_max_marks)
        
        # Check if student marks exceed total possible
        if obtained > sum(clo_max_marks):
            obtained = sum(clo_max_marks)
            
        allocated = [0.0] * len(clo_max_marks)
        remaining = obtained
        
        # Safety counter to prevent infinite loops
        iterations = 0
        while remaining > 0.001 and iterations < 5000:
            available_indices = [i for i, max_val in enumerate(clo_max_marks) if allocated[i] < max_val]
            if not available_indices: break
                
            idx = random.choice(available_indices)
            space_left = clo_max_marks[idx] - allocated[idx]
            
            # Use random increments for a more 'natural' feel (0.5 or 1.0)
            step = random.choice([0.5, 1.0])
            chunk = min(step, remaining, space_left)
            
            allocated[idx] += chunk
            remaining -= chunk
            iterations += 1
            
        return [round(m, 2) for m in allocated]
    except Exception:
        return [0.0] * len(clo_max_marks)

def main():
    st.title("🎯 OBE Marks Intelligence Pro")
    st.markdown("##### Transform raw student scores into structured CLO data and map them to your university templates.")

    # Sidebar Instructions
    with st.sidebar:
        st.header("Help & Instructions")
        st.info("""
        1. **Upload**: Support .csv, .xlsx, .xls.
        2. **Configure**: Define CLO weights.
        3. **Process**: Randomly distribute marks.
        4. **Map**: Inject data into your Excel template.
        """)
        st.divider()
        st.caption("Developed for Academic Excellence")

    tab1, tab2 = st.tabs(["📊 Distribution Engine", "📄 Template Mapper"])

    # --- TAB 1: DISTRIBUTION ---
    with tab1:
        st.header("1. Data Intake")
        uploaded_file = st.file_uploader(
            "Upload Roster", 
            type=['csv', 'xlsx', 'xls'],
            help="Upload the file containing Student Names, Roll Numbers, and Total Marks."
        )

        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                elif uploaded_file.name.endswith('.xls'):
                    df = pd.read_excel(uploaded_file, engine='xlrd')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                
                st.success(f"✅ Loaded {len(df)} records.")
                
                with st.expander("🔍 Preview Raw Data"):
                    st.dataframe(df.head(10), use_container_width=True)

                st.divider()
                st.header("2. Column Mapping & CLO Setup")
                cols = df.columns.tolist()
                
                c1, c2, c3 = st.columns(3)
                name_col = c1.selectbox("Name Column", cols, help="Select the column containing student names.")
                roll_col = c2.selectbox("Roll No Column", cols, help="Select the column containing IDs/Roll numbers.")
                marks_col = c3.selectbox("Obtained Marks", cols, help="The column with the total marks to be split.")

                # CLO Config
                num_clos = st.number_input("Total CLOs", min_value=1, value=3, help="How many CLOs are evaluated in this exam?")
                
                clo_max_marks = []
                clo_ui_cols = st.columns(num_clos)
                for i in range(num_clos):
                    m = clo_ui_cols[i].number_input(f"CLO {i+1} Max", min_value=0.1, value=10.0, key=f"m_{i}", help=f"Maximum possible marks for CLO {i+1}")
                    clo_max_marks.append(m)

                if st.button("🚀 Generate CLO Distribution"):
                    with st.spinner("Distributing marks..."):
                        # Verification
                        if not pd.api.types.is_numeric_dtype(df[marks_col]):
                            st.error(f"Error: Column '{marks_col}' contains non-numeric data. Please clean your file.")
                        else:
                            results = df[marks_col].apply(lambda x: distribute_marks(x, clo_max_marks))
                            for i in range(num_clos):
                                df[f"CLO_{i+1}"] = [res[i] for res in results]
                            
                            st.session_state['processed_df'] = df
                            st.session_state['num_clos'] = num_clos
                            st.balloons()
                            st.success("Marks Distributed! Switch to 'Template Mapper' to finalize.")
                            st.dataframe(df.head(), use_container_width=True)

            except Exception as e:
                st.error(f"Critical Error reading file: {e}")

    # --- TAB 2: MAPPING ---
    with tab2:
        if 'processed_df' not in st.session_state:
            st.warning("⚠️ Please complete the Distribution in Tab 1 first.")
        else:
            st.header("3. Official Template Integration")
            template_file = st.file_uploader(
                "Upload University Template", 
                type=['xlsx'],
                help="Upload the official empty OBE Excel sheet. It MUST be .xlsx format."
            )
            
            if template_file:
                try:
                    wb_temp = openpyxl.load_workbook(template_file)
                    sheet_name = st.selectbox("Target Sheet", wb_temp.sheetnames, help="Select the specific worksheet where marks should be written.")
                    sheet = wb_temp[sheet_name]
                    
                    st.info("📍 Map Table Coordinates")
                    m1, m2, m3 = st.columns(3)
                    start_row = m1.number_input("Start Row", min_value=1, value=2, help="The row number where the first student's data should be entered.")
                    name_target = m2.text_input("Name Col (Letter)", "B", help="Excel Column Letter for Name (e.g., B).").upper()
                    roll_target = m3.text_input("Roll Col (Letter)", "A", help="Excel Column Letter for Roll No (e.g., A).").upper()

                    st.write("---")
                    st.write("**CLO Column Mapping**")
                    clo_map_cols = st.columns(st.session_state['num_clos'])
                    clo_target_letters = []
                    for i in range(st.session_state['num_clos']):
                        let = clo_map_cols[i].text_input(f"CLO {i+1} Col", value=get_column_letter(3+i), key=f"t_{i}").upper()
                        clo_target_letters.append(let)

                    if st.button("🪄 Finalize & Export"):
                        with st.spinner("Injecting data into template..."):
                            final_df = st.session_state['processed_df']
                            for idx, row in final_df.iterrows():
                                curr = start_row + idx
                                try:
                                    sheet[f"{name_target}{curr}"] = row[name_col]
                                    sheet[f"{roll_target}{curr}"] = row[roll_col]
                                    for i in range(st.session_state['num_clos']):
                                        sheet[f"{clo_target_letters[i]}{curr}"] = row[f"CLO_{i+1}"]
                                except Exception as e:
                                    st.warning(f"Skipped row {curr} due to template error: {e}")

                            # Save
                            out_ptr = io.BytesIO()
                            wb_temp.save(out_ptr)
                            out_ptr.seek(0)
                            
                            st.download_button(
                                label="💾 Download Ready-to-Submit File",
                                data=out_ptr,
                                file_name="OBE_Final_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                except Exception as e:
                    st.error(f"Template Error: {e}")

if __name__ == "__main__":
    main()
