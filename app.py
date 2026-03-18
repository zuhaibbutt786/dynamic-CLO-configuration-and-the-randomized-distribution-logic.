import streamlit as st
import pandas as pd
import random
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OBE Marks Distributor & Mapper", layout="wide")

def distribute_marks(obtained, clo_max_marks):
    try:
        obtained = float(obtained)
    except (ValueError, TypeError):
        return [0] * len(clo_max_marks)

    if pd.isna(obtained) or obtained <= 0:
        return [0] * len(clo_max_marks)
        
    allocated = [0.0] * len(clo_max_marks)
    remaining = obtained
    
    while remaining > 0.001:
        available_indices = [i for i, max_val in enumerate(clo_max_marks) if allocated[i] < max_val]
        if not available_indices:
            break
            
        idx = random.choice(available_indices)
        space_left = clo_max_marks[idx] - allocated[idx]
        chunk = min(1.0, remaining, space_left)
        allocated[idx] += chunk
        remaining -= chunk
        
    return [round(m, 2) for m in allocated]

def main():
    st.title("🎓 OBE Marks Distributor & Template Mapper")

    # --- 1. UPLOAD DATA ---
    st.header("1. Upload Student Marks")
    uploaded_file = st.file_uploader("Upload Source File (CSV/Excel)", type=['csv', 'xlsx', 'xls'])

    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        cols = df.columns.tolist()
        
        c1, c2, c3 = st.columns(3)
        name_col = c1.selectbox("Name Column", cols)
        roll_col = c2.selectbox("Roll No Column", cols)
        marks_col = c3.selectbox("Total Obtained Marks Column", cols)

        # --- 2. CONFIGURE CLOs ---
        st.header("2. CLO Configuration")
        num_clos = st.number_input("Number of CLOs", min_value=1, value=3)
        clo_max_marks = []
        clo_ui_cols = st.columns(num_clos)
        for i in range(num_clos):
            clo_max_marks.append(clo_ui_cols[i].number_input(f"Max CLO {i+1}", min_value=1.0, value=10.0))

        # --- 3. GENERATE INTERMEDIATE DATA ---
        if st.button("Generate Randomized Splits"):
            clo_results = df[marks_col].apply(lambda x: distribute_marks(x, clo_max_marks))
            for i in range(num_clos):
                df[f"CLO_{i+1}_GEN"] = [res[i] for res in clo_results]
            st.session_state['processed_df'] = df
            st.success("Marks Distributed!")
            st.dataframe(df.head())

        # --- 4. TEMPLATE MAPPING ---
        if 'processed_df' in st.session_state:
            st.header("3. Map to OBE Template")
            template_file = st.file_uploader("Upload your Department OBE Template (Excel)", type=['xlsx'])
            
            if template_file:
                # Load template to get sheet names
                wb_temp = openpyxl.load_workbook(template_file)
                sheet_name = st.selectbox("Select Sheet in Template", wb_temp.sheetnames)
                sheet = wb_temp[sheet_name]
                
                st.info("Specify the STARTING ROW and COLUMN LETTERS in your template:")
                
                t1, t2, t3 = st.columns(3)
                start_row = t1.number_input("Starting Row (where 1st student goes)", min_value=1, value=5)
                name_target = t2.text_input("Name Column Letter (e.g., B)", "B").upper()
                roll_target = t3.text_input("Roll No Column Letter (e.g., A)", "A").upper()

                clo_target_cols = []
                st.write("Map Generated CLOs to Template Columns:")
                t_cols = st.columns(num_clos)
                for i in range(num_clos):
                    col_let = t_cols[i].text_input(f"CLO {i+1} Letter", value=get_column_letter(4+i)).upper()
                    clo_target_cols.append(col_let)

                if st.button("Finalize and Map to Template"):
                    # Writing data to the template
                    final_df = st.session_state['processed_df']
                    
                    for index, row in final_df.iterrows():
                        current_row = start_row + index
                        sheet[f"{name_target}{current_row}"] = row[name_col]
                        sheet[f"{roll_target}{current_row}"] = row[roll_col]
                        for i in range(num_clos):
                            sheet[f"{clo_target_cols[i]}{current_row}"] = row[f"CLO_{i+1}_GEN"]

                    # Save to buffer
                    out_buffer = io.BytesIO()
                    wb_temp.save(out_buffer)
                    out_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Mapped OBE Template",
                        data=out_buffer,
                        file_name="Mapped_OBE_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
