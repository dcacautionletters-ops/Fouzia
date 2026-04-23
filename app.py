import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro", layout="wide")

st.title("🎓 PG Academic Matrix: Merged Header Edition")
st.markdown("Generates a **Master Sheet** + **Sections** with merged subject headers.")

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
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
        df = df.dropna(subset=['Roll No', 'Student Name']).sort_values(by=['Section', 'Roll No'])

        # --- 4. MATRIX LOGIC ---
        def create_matrix(input_df):
            matrix = input_df.pivot_table(
                index=['Roll No', 'Student Name', 'Section'],
                columns='Course Name',
                values=['Hrs Conducted', 'Hrs Attended', 'Att %'],
                aggfunc='first'
            )
            # Course Name on TOP, Metrics BELOW
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            metrics = ['Hrs Conducted', 'Hrs Attended', 'Att %']
            matrix = matrix.reindex(columns=metrics, level=1)
            
            # Totals
            totals = input_df.groupby(['Roll No', 'Student Name', 'Section']).agg({
                'Hrs Conducted': 'sum', 'Hrs Attended': 'sum', 'Att %': 'mean'
            }).round(2)
            
            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Hrs Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Hrs Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Att %']
            
            return matrix

        master_matrix = create_matrix(df)
        st.subheader("Preview (Master Report)")
        st.dataframe(master_matrix.reset_index().fillna("-"), use_container_width=True)

        # --- 5. EXCEL EXPORT WITH ACTUAL CELL MERGING ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formats
            header_fmt = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'bg_color': '#FFCC99', 'border': 1
            })
            sub_header_fmt = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'bg_color': '#C6E0B4', 'border': 1, 'text_wrap': True
            })

            def write_sheet(matrix_data, sheet_name):
                # We reset index to make Roll/Name/Section normal columns
                flat_df = matrix_data.reset_index()
                
                # Write data starting from row 2 (to leave room for merged headers)
                flat_df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False, header=False)
                worksheet = writer.sheets[sheet_name]

                # Write Static Headers (Sl No, Roll, Name, Section)
                static_headers = ['Roll No', 'Student Name', 'Section']
                for i, col_name in enumerate(static_headers):
                    worksheet.merge_range(0, i, 1, i, col_name, header_fmt)
                
                # Write Merged Subject Headers
                # matrix_data.columns contains (Course, Metric)
                current_col = len(static_headers)
                courses = matrix_data.columns.get_level_values(0).unique()
                
                for course in courses:
                    # Merge across 3 columns (Conducted, Attended, %)
                    worksheet.merge_range(0, current_col, 0, current_col + 2, course, header_fmt)
                    
                    # Write the 3 sub-headers
                    sub_headers = ['No of Hours Conducted', 'No of Hours Attended', 'No of Attended Hours Percentage']
                    for j, sub in enumerate(sub_headers):
                        worksheet.write(1, current_col + j, sub, sub_header_fmt)
                    
                    current_col += 3

            # Write Master and Section sheets
            write_sheet(master_matrix, 'MASTER_REPORT')
            for section in sorted(df['Section'].unique()):
                section_df = df[df['Section'] == section]
                if not section_df.empty:
                    s_matrix = create_matrix(section_df)
                    write_sheet(s_matrix, str(section)[:30].replace('/', '_'))

        st.download_button(
            label="📥 Download Merged Excel Report",
            data=output.getvalue(),
            file_name="Attendance_Merged_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload file to generate the merged header report.")
