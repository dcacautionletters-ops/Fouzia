import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro", layout="wide")

st.title("🎓 PG Academic Matrix: Final Merged Edition")
st.markdown("Mapping: **B, C, G, I, J, O, P** | Version: **Multi-Sheet Fixed**")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Report", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # Load File
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
        
        # Mapping: B=1, C=2, G=6, I=8, J=9, O=14, P=15
        cols = [1, 2, 6, 8, 9, 14, 15]
        df = raw.iloc[start_row:, cols].copy()
        df.columns = ['Roll No', 'Student Name', 'Section', 'Course Name', 'Hrs Conducted', 'Hrs Attended', 'Att %']

        # --- 3. CLEANING ---
        for c in ['Hrs Conducted', 'Hrs Attended', 'Att %']:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
        df = df.dropna(subset=['Roll No', 'Student Name']).sort_values(by=['Section', 'Roll No'])

        # --- 4. MATRIX TRANSFORMATION ---
        def create_matrix(input_df):
            # Pivot the data
            matrix = input_df.pivot_table(
                index=['Roll No', 'Student Name', 'Section'],
                columns='Course Name',
                values=['Hrs Conducted', 'Hrs Attended', 'Att %'],
                aggfunc='first'
            )
            # Reorder levels: Course Name (Top), Metric (Bottom)
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            metrics_order = ['Hrs Conducted', 'Hrs Attended', 'Att %']
            matrix = matrix.reindex(columns=metrics_order, level=1)
            
            # Add Totals
            totals = input_df.groupby(['Roll No', 'Student Name', 'Section']).agg({
                'Hrs Conducted': 'sum', 'Hrs Attended': 'sum', 'Att %': 'mean'
            }).round(2)
            
            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Hrs Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Hrs Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Att %']
            return matrix

        master_matrix = create_matrix(df)
        st.subheader("Preview (Master Data)")
        st.dataframe(master_matrix.reset_index().fillna("-"), use_container_width=True)

        # --- 5. EXCEL EXPORT (The Bulletproof Fix) ---
        output = BytesIO()
        # CRITICAL: We use engine='xlsxwriter' to handle the merging manually
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Styles
            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFCC99', 'border': 1})
            sub_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'text_wrap': True})

            def write_custom_sheet(matrix_data, sheet_name):
                # STEP 1: Flatten the dataframe completely before sending to Excel
                # This prevents the 'MultiIndex columns and no index' error
                data_only = matrix_data.reset_index()
                
                # We strip the MultiIndex from the columns so Pandas sees it as a "flat" table
                data_only.columns = [f"Col_{i}" for i in range(len(data_only.columns))]
                
                # Write data starting from Row 2
                data_only.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False, header=False)
                worksheet = writer.sheets[sheet_name]

                # STEP 2: Manually Draw the Headers
                # Static Headers
                static = ['Roll No', 'Student Name', 'Section']
                for i, text in enumerate(static):
                    worksheet.merge_range(0, i, 1, i, text, header_fmt)

                # Subject Headers
                curr_col = 3
                subjects = matrix_data.columns.get_level_values(0).unique()
                for sub in subjects:
                    worksheet.merge_range(0, curr_col, 0, curr_col + 2, sub, header_fmt)
                    worksheet.write(1, curr_col, "No of Hours Conducted", sub_header_fmt)
                    worksheet.write(1, curr_col+1, "No of Hours Attended", sub_header_fmt)
                    worksheet.write(1, curr_col+2, "Att %", sub_header_fmt)
                    curr_col += 3

                # Formatting
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 1, 35)
                worksheet.set_column(2, 2, 12)
                worksheet.set_column(3, curr_col, 15)

            # Generate Master + Sections
            write_custom_sheet(master_matrix, 'MASTER_REPORT')
            for section in sorted(df['Section'].unique()):
                sect_df = df[df['Section'] == section]
                if not sect_df.empty:
                    s_matrix = create_matrix(sect_df)
                    write_custom_sheet(s_matrix, str(section)[:30].replace('/', '_'))

        st.download_button(
            label="📥 Download Professional Report",
            data=output.getvalue(),
            file_name="Attendance_Matrix_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Report Ready, Ms. Fouziya Banu!")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Awaiting file upload...")
