import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Matrix Pro", layout="wide")

st.title("🎓 PG Academic Matrix: Multi-Sheet Edition")
st.markdown("Generates a **Master Sheet** + **Individual Section Sheets** automatically.")

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
        df.columns = ['Roll No', 'Student Name', 'Section', 'Course Name', 'Conducted', 'Attended', 'Percentage']

        # --- 3. CLEANING ---
        for c in ['Conducted', 'Attended', 'Percentage']:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        df['Section'] = df['Section'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
        df = df.dropna(subset=['Roll No', 'Student Name'])
        
        # Sort by Roll No globally first
        df = df.sort_values(by='Roll No')

        # --- 4. THE MATRIX LOGIC (Function for reuse) ---
        def create_matrix(input_df):
            matrix = input_df.pivot_table(
                index=['Roll No', 'Student Name', 'Section'],
                columns='Course Name',
                values=['Conducted', 'Attended', 'Percentage'],
                aggfunc='first'
            )
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            metrics = ['Conducted', 'Attended', 'Percentage']
            matrix = matrix.reindex(columns=metrics, level=1)
            
            # Totals
            totals = input_df.groupby(['Roll No', 'Student Name', 'Section']).agg({
                'Conducted': 'sum', 'Attended': 'sum', 'Percentage': 'mean'
            }).round(2)
            
            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Percentage']
            
            final = matrix.reset_index()
            final.insert(0, 'Sl No.', range(1, len(final) + 1))
            return final

        # --- 5. GENERATE REPORTS ---
        master_matrix = create_matrix(df)
        
        st.subheader("Master Report (All Sections)")
        st.dataframe(master_matrix.fillna("-"), use_container_width=True)

        # --- 6. MULTI-SHEET EXCEL EXPORT ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # A. Save Master Sheet
            # Flatten columns for Excel compatibility
            master_export = master_matrix.copy()
            master_export.columns = [f"{c[0]} | {c[1]}".strip(' | ') if isinstance(c, tuple) else c for c in master_export.columns]
            master_export.to_excel(writer, index=False, sheet_name='MASTER_REPORT')

            # B. Save Individual Section Sheets
            unique_sections = sorted(df['Section'].unique().tolist())
            for section in unique_sections:
                section_df = df[df['Section'] == section]
                if not section_df.empty:
                    sect_matrix = create_matrix(section_df)
                    # Flatten for Excel
                    sect_matrix.columns = [f"{c[0]} | {c[1]}".strip(' | ') if isinstance(c, tuple) else c for c in sect_matrix.columns]
                    # Clean sheet name (max 31 chars, no special chars)
                    sheet_name = str(section)[:30].replace('/', '_')
                    sect_matrix.to_excel(writer, index=False, sheet_name=sheet_name)

            # Formatting
            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            for sheet in writer.sheets.values():
                sheet.set_column(0, 50, 18)

        st.download_button(
            label="📥 Download Multi-Sheet Excel (Master + Sections)",
            data=output.getvalue(),
            file_name="Attendance_Matrix_Split_By_Section.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload your raw file to generate the multi-sheet matrix.")
