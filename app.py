import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro", layout="wide")

st.title("🎓 PG Academic Matrix: Final Custom Edition")
st.markdown("Mapping: **B, C, G, I, J, O, P** | Sl No, Alignments & Shrinkage Active")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Report", type=['csv', 'xlsx'])

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

        # --- 3. CLEANING ---
        for c in ['Hrs Conducted', 'Hrs Attended', 'Att %']:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        
        df = df.replace([np.inf, -np.inf], np.nan).fillna(0)
        df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
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
        st.subheader("Final Matrix Preview")
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
                # 1. Flatten Data and add Sl No
                flat_df = matrix_data.reset_index()
                flat_df.insert(0, 'Sl No.', range(1, len(flat_df) + 1))
                
                total_rows, total_cols = flat_df.shape
                # Dummy columns to bypass pandas MultiIndex logic
                flat_df.columns = [f"Col_{i}" for i in range(total_cols)]
                
                flat_df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False, header=False)
                worksheet = writer.sheets[sheet_name]

                # 2. Apply Formatting Row-by-Row
                for r in range(total_rows):
                    for c in range(total_cols):
                        val = flat_df.iloc[r, c]
                        # Student Name is at index 2 (Sl No=0, Roll=1, Name=2)
                        if c == 2:
                            worksheet.write(r + 2, c, val, left_data_fmt)
                        else:
                            worksheet.write(r + 2, c, val, center_data_fmt)

                # 3. Draw Headers
                static = ['Sl No.', 'Roll No', 'Student Name', 'Section']
                for i, text in enumerate(static):
                    worksheet.merge_range(0, i, 1, i, text, header_fmt)

                curr_col = 4 # Starting after Section
                subjects = matrix_data.columns.get_level_values(0).unique()
                for sub in subjects:
                    worksheet.merge_range(0, curr_col, 0, curr_col + 2, sub, header_fmt)
                    worksheet.write(1, curr_col, "Hrs Cond.", sub_header_fmt)
                    worksheet.write(1, curr_col+1, "Hrs Attd.", sub_header_fmt)
                    worksheet.write(1, curr_col+2, "Att %", sub_header_fmt)
                    curr_col += 3

                # 4. Shrink Column Widths
                worksheet.set_column(0, 0, 6)   # Sl No
                worksheet.set_column(1, 1, 15)  # Roll
                worksheet.set_column(2, 2, 35)  # Name (Keep wide for visibility)
                worksheet.set_column(3, 3, 12)  # Section
                worksheet.set_column(4, curr_col, 10) # Metrics (Shrunken)

            # Generate sheets
            write_custom_sheet(master_matrix, 'MASTER_REPORT')
            for section in sorted(df['Section'].unique()):
                sect_df = df[df['Section'] == section]
                if not sect_df.empty:
                    write_custom_sheet(create_matrix(sect_df), str(section)[:30].replace('/', '_'))

        st.download_button(
            label="📥 Download Final Customized Report",
            data=output.getvalue(),
            file_name="Final_Attendance_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("All alignments and Sl No. columns are ready! Enjoy your day Fouziya!")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Awaiting file upload...")
