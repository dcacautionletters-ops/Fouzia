import streamlit as st
import pandas as pd
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="Universal PG Matrix", layout="wide")

st.title("🎓 Smart PG Academic Matrix")
st.markdown("Dynamic detection of columns: **B, C, G, I, J, O, P**")

# --- 1. FILE UPLOAD ---
uploaded_file = st.file_uploader("Upload Raw Report", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        # Load the whole sheet first to scan for headers
        if uploaded_file.name.endswith('.csv'):
            full_df = pd.read_csv(uploaded_file, header=None)
        else:
            full_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

        # --- 2. DYNAMIC HEADER DETECTION ---
        # Find the row where "Roll" or "Student" exists in the first 2 columns
        start_row = 0
        for r in range(20):
            row_values = [str(val).lower() for val in full_df.iloc[r, 0:3]]
            if any("roll" in v or "reg" in v for v in row_values):
                start_row = r + 1 # Data starts after header
                break
        
        # Mapping: B=1, C=2, G=6, I=8, J=9, O=14, P=15
        # We use .iloc to be index-perfect regardless of row start
        data_cols = [1, 2, 6, 8, 9, 14, 15]
        df = full_df.iloc[start_row:, data_cols].copy()
        df.columns = ['Roll No', 'Student Name', 'Batch', 'Course Name', 'Conducted', 'Attended', 'Percentage']

        # --- 3. CLEANING & "STR VS FLOAT" FIX ---
        # 1. Clean up numeric columns
        for col in ['Conducted', 'Attended', 'Percentage']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 2. Fix the sorting error: Force Batch and Course to be strings
        df['Batch'] = df['Batch'].astype(str).replace('nan', 'Unknown').str.strip()
        df['Course Name'] = df['Course Name'].astype(str).replace('nan', 'Unknown').str.strip()
        
        # 3. Drop completely empty rows
        df = df.dropna(subset=['Roll No', 'Student Name'])

        # --- 4. SIDEBAR FILTERS ---
        all_batches = sorted(df['Batch'].unique().tolist())
        selected_batches = st.sidebar.multiselect("Filter by Batch", all_batches, default=all_batches)
        
        df_filtered = df[df['Batch'].isin(selected_batches)]

        if not df_filtered.empty:
            # --- 5. MATRIX TRANSFORMATION ---
            matrix = df_filtered.pivot_table(
                index=['Roll No', 'Student Name', 'Batch'],
                columns='Course Name',
                values=['Conducted', 'Attended', 'Percentage'],
                aggfunc='first'
            )

            # Format: Subject > Metrics
            matrix = matrix.reorder_levels([1, 0], axis=1).sort_index(axis=1)
            metrics = ['Conducted', 'Attended', 'Percentage']
            matrix = matrix.reindex(columns=metrics, level=1)

            # --- 6. TOTALS ---
            totals = df_filtered.groupby(['Roll No', 'Student Name', 'Batch']).agg({
                'Conducted': 'sum',
                'Attended': 'sum',
                'Percentage': 'mean'
            }).round(2)

            matrix[('GRAND TOTAL', 'Total Conducted')] = totals['Conducted']
            matrix[('GRAND TOTAL', 'Total Attended')] = totals['Attended']
            matrix[('GRAND TOTAL', 'Average %')] = totals['Percentage']

            # --- 7. DISPLAY & EXPORT ---
            final_df = matrix.reset_index()
            final_df.insert(0, 'Sl No.', range(1, len(final_df) + 1))
            
            st.subheader("Attendance Overview")
            st.dataframe(final_df.fillna("-"), use_container_width=True)

            # Flatten headers for Excel download
            export_df = final_df.copy()
            if isinstance(export_df.columns, pd.MultiIndex):
                export_df.columns = [f"{c[0]} - {c[1]}".strip(' - ') for c in export_df.columns]

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Matrix')
            
            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name="Final_Attendance_Matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data found. Please check your filters.")

    except Exception as e:
        st.error(f"Critical Error: {e}")
        st.info("Check if the file has data in columns B, C, G, I, J, O, P.")
