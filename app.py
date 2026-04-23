import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro", layout="wide")

# --- SECURITY LOCK ---
def check_password():
    """Returns True if the user had the correct password."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == "VMS@123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Enter Password to Access PG Matrix Pro", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect, show input + error.
        st.text_input(
            "Enter Password to Access PG Matrix Pro", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if check_password():
    # --- ORIGINAL APP CODE STARTS HERE ---
    st.title("🎓 PG Academic Matrix: Universal Edition")
    st.markdown("Mapping: **B, C, G, I, J, O, P** | **Left-Aligned Names & Blacklisted Free Slots**")

    # --- 1. FILE UPLOAD ---
    uploaded_file = st.file_uploader("Upload Consolidated or Separate Reports", type=['csv', 'xlsx'])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                raw = pd.read_csv(uploaded_file, header=None)
            else:
                raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

            # --- 2. DYNAMIC DETECTION ---
            start_row = 0
            for r in range(20):
                row_vals = [str(val).lower() for val in raw.iloc[r, 0:5]]
                if any("roll" in v or "reg" in v for v in row_vals):
                    start_row = r + 1 
                    break
            
            cols = [1, 2, 6, 8, 9, 14, 15]
            df = raw.iloc[start_row:, cols].copy()
            df.columns = ['Roll No', 'Student Name', 'Section', 'Course Name', 'Hrs Conducted', 'Hrs Attended', 'Att %']

            # --- 3. CLEANING & BLACKLISTING ---
            df['Course Name'] = df['Course Name'].astype(str).str.strip()
            blacklist_keywords = ['freeslot', 'free slot']
            df = df[~df['Course Name'].str.lower().str.replace(' ', '').isin(['freeslot'])]
            df = df[~df['Course Name'].str.lower().isin(blacklist_keywords)]

            for c in ['Hrs Conducted', 'Hrs Attended', 'Att %']:
                df[c] = pd.to_numeric(df[c], errors='coerce')
            
            df = df.replace([np.inf, -np.inf], np.nan).fillna(0)
            df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
            df = df.dropna(subset=['Roll No', 'Student Name']).sort_values(by=['Section', 'Roll No'])

            # --- 4. MATRIX TRANSFORMATION ---
            def create_matrix(input_df):
                matrix = input_df.pivot_table(
                    index=['Roll No', 'Student Name', 'Section'],
                    columns='Course Name',
                    values=['Hrs Conducted', 'Hrs Attended', 'Att %'],
                    aggfunc='first'
                )
                matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
                metrics_order = ['Hrs Conducted', 'Hrs Attended', 'Att %']
                matrix = matrix.reindex(columns=metrics_order, level=1)
                
                totals = input_df.groupby(['Roll No', 'Student Name', 'Section']).agg({
                    'Hrs Conducted': 'sum', 'Hrs Attended': 'sum', 'Att %': 'mean'
                }).round(2)
                
                matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Hrs Conducted']
                matrix[('GRAND TOTAL', 'Total Attended')] = totals['Hrs Attended']
                matrix[('GRAND TOTAL', 'Average %')] = totals['Att %']
                return matrix.fillna(0)

            master_matrix = create_matrix(df)
            st.subheader("Global Preview (Consolidated)")
            st.dataframe(master_matrix.reset_index().fillna("-"), use_container_width=True)

            # --- 5. EXCEL EXPORT (CUSTOM ALIGNMENT & BORDERS) ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                workbook = writer.book
                
                # --- FORMATS ---
                header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFCC99', 'border': 1, 'text_wrap': True})
                sub_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'text_wrap': True})
                
                # Left Align for Names
                left_data_fmt = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
                # Center Align for Numbers
                center_data_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

                def write_custom_sheet(matrix_data, sheet_name):
                    # Flatten and add Sl No
                    flat_df = matrix_data.reset_index()
                    flat_df.insert(0, 'Sl No.', range(1, len(flat_df) + 1))
                    
                    total_rows, total_cols = flat_df.shape
                    # Rename columns for flat export
                    flat_df.columns = [f"Col_{i}" for i in range(total_cols)]
                    
                    flat_df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False, header=False)
                    worksheet = writer.sheets[sheet_name]

                    # --- ROW-BY-ROW FORMATTING ---
                    for r in range(total_rows):
                        for c in range(total_cols):
                            val = flat_df.iloc[r, c]
                            # 0: Sl No, 1: Roll No, 2: Student Name, 3: Section
                            if c == 2: # Student Name is definitely at index 2
                                worksheet.write(r + 2, c, val, left_data_fmt)
                            else:
                                worksheet.write(r + 2, c, val, center_data_fmt)

                    # --- DRAW HEADERS ---
                    static = ['Sl No.', 'Roll No', 'Student Name', 'Section']
                    for i, text in enumerate(static):
                        worksheet.merge_range(0, i, 1, i, text, header_fmt)

                    curr_col = 4
                    subjects = matrix_data.columns.get_level_values(0).unique()
                    for sub in subjects:
                        worksheet.merge_range(0, curr_col, 0, curr_col + 2, sub, header_fmt)
                        worksheet.write(1, curr_col, "Hrs Cond.", sub_header_fmt)
                        worksheet.write(1, curr_col+1, "Hrs Attd.", sub_header_fmt)
                        worksheet.write(1, curr_col+2, "Att %", sub_header_fmt)
                        curr_col += 3

                    # --- COLUMN WIDTHS ---
                    worksheet.set_column(0, 0, 6)   # Sl No
                    worksheet.set_column(1, 1, 15)  # Roll
                    worksheet.set_column(2, 2, 35)  # Name (Left Aligned)
                    worksheet.set_column(3, 3, 12)  # Section
                    worksheet.set_column(4, curr_col, 10) # Subjects & Totals (Shrunken)

                # Generate Master Sheet
                write_custom_sheet(master_matrix, 'MASTER_REPORT')
                
                # Automatically generate individual tabs for MCA, MBA, etc.
                for section in sorted(df['Section'].unique()):
                    sect_df = df[df['Section'] == section]
                    if not sect_df.empty:
                        write_custom_sheet(create_matrix(sect_df), str(section)[:31].replace('/', '_'))

            st.download_button(
                label="📥 Download Universal Report",
                data=output.getvalue(),
                file_name="Final_Universal_Attendance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Everything is set! Left-aligned names and blacklisted free slots are active.")

        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.info("Awaiting file upload...")
