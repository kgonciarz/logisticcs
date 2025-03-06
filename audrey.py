import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.title("Excel Processing App (Preserve Formulas)")

# File uploaders for both received file and template
uploaded_file = st.file_uploader("Upload the received Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload the template Excel file", type=["xlsx"])

if uploaded_file and template_file:
    # Load the uploaded file into a Pandas DataFrame (read first sheet without headers)
    df_uploaded = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Load the template file with openpyxl
    wb = load_workbook(template_file)
    ws_feuil1 = wb["Feuil1"]  # Ensure we are working with "Feuil1"

    # ✅ Step 1: Clear Old Data but Keep Formulas
    for row in ws_feuil1.iter_rows(min_row=2, max_row=ws_feuil1.max_row, min_col=1, max_col=ws_feuil1.max_column):
        for cell in row:
            if not cell.data_type == "f":  # Only clear non-formula cells
                cell.value = None

    # ✅ Step 2: Copy and Paste New Data Without Overwriting Formulas
    for i, row in enumerate(df_uploaded.values, start=2):  # Start at row 2 to keep header row
        for j, value in enumerate(row, start=1):
            ws_feuil1.cell(row=i, column=j, value=value)

    # ✅ Step 3: Save the updated workbook to a BytesIO buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # ✅ Step 4: Provide Download Button
    st.download_button(
        label="📥 Download Processed File with Formulas",
        data=output,
        file_name="Processed_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
