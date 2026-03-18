import streamlit as st
import pandas as pd
import random
import io

st.set_page_config(page_title="CLO Marks Distributor", layout="wide")

def distribute_marks(obtained, clo_max_marks):
    """
    Randomly distributes the obtained marks across CLOs without exceeding max marks.
    Handles both integers and standard float increments (e.g., 0.5, 1.0).
    """
    try:
        obtained = float(obtained)
    except (ValueError, TypeError):
        return [0] * len(clo_max_marks)

    if pd.isna(obtained) or obtained <= 0:
        return [0] * len(clo_max_marks)
        
    allocated = [0.0] * len(clo_max_marks)
    remaining = obtained
    
    # Loop until all marks are distributed (using a small threshold for float precision)
    while remaining > 0.001:
        # Find which CLOs can still accept marks
        available_indices = [i for i, max_val in enumerate(clo_max_marks) if allocated[i] < max_val]
        
        if not available_indices:
            break # Student obtained marks exceed total available CLO marks
            
        idx = random.choice(available_indices)
        
        # Calculate how much space is left in this specific CLO
        space_left = clo_max_marks[idx] - allocated[idx]
        
        # Add marks in chunks of 1 (or the remaining fraction if less than 1)
        # This simulates realistic grading steps
        chunk = min(1.0, remaining, space_left)
        
        allocated[idx] += chunk
        remaining -= chunk
        
    # Round to avoid weird floating point artifacts (e.g., 4.999999)
    return [round(m, 2) for m in allocated]

def main():
    st.title("🎓 CLO Marks Random Distributor")
    st.markdown("Upload a student roster, define your CLOs, and randomly split their total marks into individual CLO scores.")

    # --- STEP 1: FILE UPLOAD ---
    st.header("1. Upload Data")
    uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=['csv', 'xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success("File uploaded successfully!")
            with st.expander("Preview Uploaded Data"):
                st.dataframe(df.head())
                
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

        # --- STEP 2: COLUMN MAPPING ---
        st.header("2. Map Columns")
        cols = df.columns.tolist()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            name_col = st.selectbox("Select 'Name' Column", options=["Skip"] + cols)
        with col2:
            roll_col = st.selectbox("Select 'Roll No' Column", options=["Skip"] + cols)
        with col3:
            marks_col = st.selectbox("Select 'Obtained Marks' Column", options=cols, index=len(cols)-1)

        # --- STEP 3: CLO CONFIGURATION ---
        st.header("3. Configure CLOs")
        num_clos = st.number_input("How many CLOs are there?", min_value=1, max_value=20, value=3)
        
        clo_max_marks = []
        clo_cols = st.columns(num_clos)
        
        for i in range(num_clos):
            with clo_cols[i]:
                max_mark = st.number_input(f"Max Marks for CLO {i+1}", min_value=1.0, value=10.0, step=1.0)
                clo_max_marks.append(max_mark)
                
        total_clo_marks = sum(clo_max_marks)
        st.info(f"**Total Available Marks (Sum of CLOs): {total_clo_marks}**")

        # --- STEP 4: PROCESSING ---
        st.header("4. Process and Download")
        if st.button("Generate CLO Split", type="primary"):
            # Check if any student has more marks than the max allowed
            max_student_marks = df[marks_col].max()
            if max_student_marks > total_clo_marks:
                st.warning(f"Warning: A student has {max_student_marks} marks, which is higher than the total CLO marks ({total_clo_marks}). Their marks will be capped.")

            # Create a copy for the output
            out_df = df.copy()
            
            # Apply the distribution logic
            clo_results = out_df[marks_col].apply(lambda x: distribute_marks(x, clo_max_marks))
            
            # Create new columns for each CLO
            for i in range(num_clos):
                clo_name = f"CLO {i+1} (Max: {clo_max_marks[i]})"
                out_df[clo_name] = [res[i] for res in clo_results]
                
            st.success("Marks distributed successfully!")
            st.dataframe(out_df)

            # Export logic
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                out_df.to_excel(writer, index=False, sheet_name='CLO Marks')
            output.seek(0)

            st.download_button(
                label="📥 Download Processed Excel File",
                data=output,
                file_name="distributed_clo_marks.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
