import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="PG Attendance Matrix", layout="wide")

st.title("📊 PG Academic Matrix Reporter")
st.markdown("Mapping: B(Roll), C(Name), G(Batch), I(Course), J(Cond.), O(Attd.), P(%)")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Attendance Report", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # Load the file without headers initially to find the data row
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file, header=None)
        else:
            raw_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

        # --- 2. FLEXIBLE HEADER DETECTION ---
        # Scan first few rows to find where data starts (usually where 'Roll' is in Col B/Index 1)
        start_row = 0
        for i in range(10):
            val = str(raw_df.iloc[i, 1]).lower()
            if "roll" in val or "2" in val: # common starting points for roll numbers
                start_row = i
                break
        
        # Column Mapping (B=1, C=2, G=6, I=8, J=9, O=14, P=15)
        cols = [1, 2, 6, 8, 9, 14, 15]
        df = raw_df.iloc[start_row:, cols].copy()
        df.columns = ['Roll No', 'Student Name', 'Batch', 'Course Name', 'Conducted', 'Attended', 'Percentage']
        
        # Convert numbers and cleanup
        df['Conducted'] = pd.to_numeric(df['Conducted'], errors='coerce')
        df['Attended'] = pd.to_numeric(df['Attended'], errors='coerce')
        df['Percentage'] = pd.to_numeric(df['Percentage'], errors='coerce')
        df = df.dropna(subset=['Roll No', 'Student Name'])

        # --- 3. SIDEBAR FILTERS ---
        batches = sorted(df['Batch'].unique().astype(str).tolist())
        selected_batches = st.sidebar.multiselect("Select Batches", batches, default=batches)
        df_filtered = df[df['Batch'].astype(str).isin(selected_batches)]

        if not df_filtered.empty:
            # --- 4. MATRIX TRANSFORMATION ---
            # Pivot to get Course Names as horizontal headers
            matrix = df_filtered.pivot_table(
                index=['Roll No', 'Student Name', 'Batch'],
                columns='Course Name',
                values=['Conducted', 'Attended', 'Percentage'],
                aggfunc='first'
            )

            # Reorder levels: [Course Name] on top, [Conducted/Attended/%] below
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            metric_order = ['Conducted', 'Attended', 'Percentage']
            matrix = matrix.reindex(columns=metric_order, level=1)

            # --- 5. CALCULATE GRAND TOTALS ---
            totals = df_filtered.groupby(['Roll No', 'Student Name', 'Batch']).agg({
                'Conducted': 'sum',
                'Attended': 'sum',
                'Percentage': 'mean'
            }).round(2)

            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Percentage']

            # --- 6. DISPLAY ---
            final_df = matrix.reset_index()
            final_df.insert(0, 'Sl No.', range(1, len(final_df) + 1))
            st.subheader("Attendance Matrix View")
            st.dataframe(final_df.fillna("-"), use_container_width=True)

            # --- 7. EXCEL EXPORT (Fix for MultiIndex Error) ---
            output = BytesIO()
            export_df = final_df.copy()

            # Flatten headers for Excel compatibility
            if isinstance(export_df.columns, pd.MultiIndex):
                export_df.columns = [
                    f"{col[0]}_{col[1]}".strip('_') if isinstance(col, tuple) else col 
                    for col in export_df.columns.values
                ]

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Report')
                
                # Format adjustment
                worksheet = writer.sheets['Report']
                for idx, col in enumerate(export_df.columns):
                    max_len = max(export_df[col].astype(str).map(len).max(), len(str(col))) + 2
                    worksheet.set_column(idx, idx, min(max_len, 40))

            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name="Attendance_Matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data found for the selected batches.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("At your service Madam Ji! Jiyo Jithe Raho!!.")