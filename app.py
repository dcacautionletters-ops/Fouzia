import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Attendance Matrix", layout="wide")

st.title("🎓 PG Academic Matrix Generator")
st.markdown("Groups metrics (Hours, Attended, %) under each **Course Name**.")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Report", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # Load the file
        if uploaded_file.name.endswith('.csv'):
            full_df = pd.read_csv(uploaded_file, header=None)
        else:
            full_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

        # --- 2. FLEXIBLE HEADER DETECTION ---
        # Scans first 20 rows to find where the data starts
        start_row = 0
        for r in range(20):
            row_values = [str(val).lower() for val in full_df.iloc[r, 0:5]]
            if any("roll" in v or "reg" in v for v in row_values):
                start_row = r + 1 
                break
        
        # Mapping: B=1, C=2, G=6, I=8, J=9, O=14, P=15
        data_cols = [1, 2, 6, 8, 9, 14, 15]
        df = full_df.iloc[start_row:, data_cols].copy()
        
        # Use descriptive names for the grouping logic
        df.columns = [
            'Roll No', 'Student Name', 'Batch', 'Course Name', 
            'Hrs Conducted', 'Hrs Attended', 'Att %'
        ]

        # --- 3. DATA CLEANING (Prevents str vs float error) ---
        for col in ['Hrs Conducted', 'Hrs Attended', 'Att %']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Force strings to avoid sorting crashes
        df['Batch'] = df['Batch'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).str.strip()
        df = df.dropna(subset=['Roll No', 'Student Name'])

        # --- 4. SIDEBAR FILTERS ---
        all_batches = sorted(df['Batch'].unique().tolist())
        selected_batches = st.sidebar.multiselect("Filter by Batch", all_batches, default=all_batches)
        df_filtered = df[df['Batch'].isin(selected_batches)]

        if not df_filtered.empty:
            # --- 5. THE "CLUBBED" PIVOT TABLE ---
            # This creates the hierarchy: Course Name > Metrics
            matrix = df_filtered.pivot_table(
                index=['Roll No', 'Student Name', 'Batch'],
                columns='Course Name',
                values=['Hrs Conducted', 'Hrs Attended', 'Att %'],
                aggfunc='first'
            )

            # SWAP levels so the Course Name is the TOP header
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            
            # Ensure the 3 columns under each course are in the specific order you asked for
            metrics = ['Hrs Conducted', 'Hrs Attended', 'Att %']
            matrix = matrix.reindex(columns=metrics, level=1)

            # --- 6. GRAND TOTALS ---
            totals = df_filtered.groupby(['Roll No', 'Student Name', 'Batch']).agg({
                'Hrs Conducted': 'sum',
                'Hrs Attended': 'sum',
                'Att %': 'mean'
            }).round(2)

            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Hrs Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Hrs Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Att %']

            # --- 7. DISPLAY ---
            # Reset index and add Serial Number
            final_df = matrix.reset_index()
            final_df.insert(0, 'Sl No.', range(1, len(final_df) + 1))
            
            st.subheader("Subject-Wise Attendance Matrix")
            # This displays the "Clubbed" headers in Streamlit
            st.dataframe(final_df.fillna("-"), use_container_width=True)

            # --- 8. EXCEL EXPORT (Cleaning for Download) ---
            output = BytesIO()
            export_df = final_df.copy()
            
            # Excel doesn't always handle MultiIndex 'index=False' well, 
            # so we flatten just for the download file to prevent errors.
            if isinstance(export_df.columns, pd.MultiIndex):
                export_df.columns = [
                    f"{c[0]} | {c[1]}".strip(' | ') if c[1] else c[0] 
                    for c in export_df.columns
                ]

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Matrix')
                
                # Auto-format column width
                worksheet = writer.sheets['Matrix']
                for idx, col in enumerate(export_df.columns):
                    worksheet.set_column(idx, idx, 20)

            st.download_button(
                label="📥 Download Grouped Excel Report",
                data=output.getvalue(),
                file_name="Grouped_Attendance_Matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data found for selected filters.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check if Column G (Batch) or Column I (Course Name) contains valid data.")
else:
    st.info("Upload your raw report to generate the grouped matrix.")
