import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.title("Excel Processing App")

# File uploaders for both received file and template
uploaded_file = st.file_uploader("Upload the received Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload the template Excel file", type=["xlsx"])

if uploaded_file and template_file:
    # Load the uploaded file and template
    df_uploaded = pd.read_excel(uploaded_file, sheet_name=0, header=None)  # Read first sheet without headers
    wb = load_workbook(template_file)
    ws_feuil1 = wb["Feuil1"]

    # Clear all existing data in Feuil1
    for row in ws_feuil1.iter_rows(min_row=1, max_row=ws_feuil1.max_row, min_col=1, max_col=ws_feuil1.max_column):
        for cell in row:
            cell.value = None

    # Copy and paste all data from the uploaded file into Feuil1
    for i, row in enumerate(df_uploaded.values, start=1):
        for j, value in enumerate(row, start=1):
            ws_feuil1.cell(row=i, column=j, value=value)

    # Save the updated workbook to a BytesIO buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # Provide download option with formulas preserved
    st.download_button(
        label="Download Processed File with Formulas",
        data=output,
        file_name="Processed_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
