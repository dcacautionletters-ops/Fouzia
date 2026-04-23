import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Attendance Matrix", layout="wide")

# Custom Title
st.title("🎓 PG Academic Matrix Generator")
st.markdown("Mapping: **B**(Roll), **C**(Name), **G**(Batch), **I**(Course), **J**(Cond.), **O**(Attd.), **P**(%)")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Attendance Report (Excel or CSV)", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # Load the file
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file, header=None)
        else:
            raw_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

        # --- 2. FLEXIBLE HEADER DETECTION ---
        # Look for the start of data (Scanning first 10 rows for 'Roll' or row with data)
        start_row = 0
        for i in range(10):
            val = str(raw_df.iloc[i, 1]).lower() # Checking Column B (Index 1)
            if "roll" in val or (val != 'nan' and len(val) > 1):
                start_row = i
                # If the current row is the header (contains 'Roll'), skip it to get to data
                if "roll" in val:
                    start_row = i + 1
                break
        
        # Column Mapping based on your input:
        # B=1, C=2, G=6, I=8, J=9, O=14, P=15
        cols = [1, 2, 6, 8, 9, 14, 15]
        df = raw_df.iloc[start_row:, cols].copy()
        df.columns = ['Roll No', 'Student Name', 'Batch', 'Course Name', 'Conducted', 'Attended', 'Percentage']
        
        # --- 3. DATA CLEANING ---
        # Force numeric types for calculations
        df['Conducted'] = pd.to_numeric(df['Conducted'], errors='coerce')
        df['Attended'] = pd.to_numeric(df['Attended'], errors='coerce')
        df['Percentage'] = pd.to_numeric(df['Percentage'], errors='coerce')
        
        # Convert Batch to string to avoid "str vs float" sorting errors
        df['Batch'] = df['Batch'].astype(str).replace('nan', 'Unknown')
        
        # Drop rows where Name or Roll No is missing
        df = df.dropna(subset=['Roll No', 'Student Name'])

        # --- 4. SIDEBAR FILTERS ---
        batches = sorted(df['Batch'].unique().tolist())
        selected_batches = st.sidebar.multiselect("Select Batches", batches, default=batches)
        df_filtered = df[df['Batch'].isin(selected_batches)]

        if not df_filtered.empty:
            # --- 5. MATRIX TRANSFORMATION ---
            # Create the pivot table (Subjects across the top)
            matrix = df_filtered.pivot_table(
                index=['Roll No', 'Student Name', 'Batch'],
                columns='Course Name',
                values=['Conducted', 'Attended', 'Percentage'],
                aggfunc='first'
            )

            # Fix Hierarchy: [Course Name] -> [Metrics]
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            
            # Ensure order of Conducted, Attended, % for every subject
            metric_order = ['Conducted', 'Attended', 'Percentage']
            matrix = matrix.reindex(columns=metric_order, level=1)

            # --- 6. GRAND TOTALS ---
            totals = df_filtered.groupby(['Roll No', 'Student Name', 'Batch']).agg({
                'Conducted': 'sum',
                'Attended': 'sum',
                'Percentage': 'mean'
            }).round(2)

            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Percentage']

            # --- 7. UI DISPLAY ---
            final_df = matrix.reset_index()
            final_df.insert(0, 'Sl No.', range(1, len(final_df) + 1))
            
            st.subheader(f"Generated Matrix ({len(final_df)} Students)")
            st.dataframe(final_df.fillna("-"), use_container_width=True)

            # --- 8. EXCEL EXPORT (FLATTENED HEADERS) ---
            output = BytesIO()
            export_df = final_df.copy()

            # Flatten MultiIndex columns to avoid Excel Export error
            if isinstance(export_df.columns, pd.MultiIndex):
                export_df.columns = [
                    f"{col[0]} - {col[1]}".strip(' - ') if isinstance(col, tuple) else col 
                    for col in export_df.columns.values
                ]

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Attendance_Report')
                
                # Basic formatting for the Excel file
                workbook = writer.book
                worksheet = writer.sheets['Attendance_Report']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
                
                for idx, col in enumerate(export_df.columns):
                    worksheet.set_column(idx, idx, 18)

            st.download_button(
                label="📥 Download Matrix as Excel",
                data=output.getvalue(),
                file_name="Attendance_Matrix_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Please select at least one batch from the sidebar.")

    except Exception as e:
        st.error(f"Error Processing File: {e}")
        st.info("Ensure your file has data starting near Row 3/4 and columns match B, C, G, I, J, O, P.")
else:
    st.info("👋 Ready! Please upload your raw Excel/CSV file to begin.")
