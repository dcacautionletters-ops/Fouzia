import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro: Multi-Series", layout="wide")

st.title("🎓 PG Academic Matrix: Weighted Average Edition")
st.markdown("Ensuring **Grand Total %** is calculated as (Total Attended / Total Conducted).")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Consolidated Reports", type=['csv', 'xlsx'])

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

        # --- 3. CLEANING & BATCH IDENTIFICATION ---
        def detect_batch(row):
            val = (str(row['Roll No']) + str(row['Section']) + str(row['Course Name'])).upper()
            program = "MCA" if "MCA" in val else "MBA" if "MBA" in val else "MFA" if "MFA" in val else "PG"
            year = "2025" if "25" in val else "2024" if "24" in val else ""
            return f"{program} {year}".strip()

        df['Batch'] = df.apply(detect_batch, axis=1)
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
        
        # Blacklisting
        blacklist_keywords = ['freeslot', 'free slot']
        df = df[~df['Course Name'].str.lower().str.replace(' ', '').isin(['freeslot'])]
        df = df[~df['Course Name'].str.lower().isin(blacklist_keywords)]

        for c in ['Hrs Conducted', 'Hrs Attended']:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        
        df = df.replace([np.inf, -np.inf], np.nan).fillna(0)
        df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
        df = df.dropna(subset=['Roll No', 'Student Name']).sort_values(by=['Batch', 'Section', 'Roll No'])

        # --- 4. MATRIX TRANSFORMATION (WITH WEIGHTED AVERAGE) ---
        def create_matrix(input_df):
            # Pivot the main data
            matrix = input_df.pivot_table(
                index=['Roll No', 'Student Name', 'Batch', 'Section'],
                columns='Course Name',
                values=['Hrs Conducted', 'Hrs Attended', 'Att %'],
                aggfunc='first'
            )
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            
            # Calculate Totals
            totals = input_df.groupby(['Roll No', 'Student Name', 'Batch', 'Section']).agg({
                'Hrs Conducted': 'sum', 
                'Hrs Attended': 'sum'
            })
            
            # THE FIX: Weighted Average Calculation
            # We avoid 0 division by adding a tiny epsilon or using np.where
            totals['Weighted %'] = np.where(
                totals['Hrs Conducted'] > 0, 
                (totals['Hrs Attended'] / totals['Hrs Conducted']) * 100, 
                0
            ).round(2)
            
            # Append Grand Totals to Matrix
            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Hrs Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Hrs Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Weighted %']
            
            return matrix.fillna(0)

        # --- 5. EXCEL EXPORT (STYLING) ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
            sub_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1})
            left_fmt = workbook.add_format({'align': 'left', 'border': 1})
            center_fmt = workbook.add_format({'align': 'center', 'border': 1})

            def write_sheet(matrix_data, sheet_name):
                flat_df = matrix_data.reset_index()
                flat_df.insert(0, 'Sl No.', range(1, len(flat_df) + 1))
                total_rows, total_cols = flat_df.shape
                
                ws = workbook.add_worksheet(sheet_name[:31])
                # Write data
                for r in range(total_rows):
                    for c in range(total_cols):
                        val = flat_df.iloc[r, c]
                        ws.write(r + 2, c, val, left_fmt if c == 2 else center_fmt)

                # Headers
                static_cols = ['Sl No.', 'Roll No', 'Student Name', 'Batch', 'Section']
                for i, text in enumerate(static_cols):
                    ws.merge_range(0, i, 1, i, text, header_fmt)

                curr_col = 5
                subjects = matrix_data.columns.get_level_values(0).unique()
                for sub in subjects:
                    ws.merge_range(0, curr_col, 0, curr_col + 2, sub, header_fmt)
                    ws.write(1, curr_col, "Cond.", sub_header_fmt)
                    ws.write(1, curr_col+1, "Attd.", sub_header_fmt)
                    ws.write(1, curr_col+2, "%", sub_header_fmt)
                    curr_col += 3
                
                ws.set_column(2, 2, 35) # Student Name width

            # Generate Sheets
            master_matrix = create_matrix(df)
            write_sheet(master_matrix, 'MASTER_REPORT')

            for batch in sorted(df['Batch'].unique()):
                batch_df = df[df['Batch'] == batch]
                write_sheet(create_matrix(batch_df), f"{batch}_TOTAL")
                
                for section in sorted(batch_df['Section'].unique()):
                    sect_df = batch_df[batch_df['Section'] == section]
                    write_sheet(create_matrix(sect_df), f"{batch}_{section}")

        st.download_button(
            label="📥 Download Corrected Reports",
            data=output.getvalue(),
            file_name="PG_Attendance_Corrected_Averages.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Mathematical correction applied! Average % is now based on total hours.")

    except Exception as e:
        st.error(f"Error: {e}")
