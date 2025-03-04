import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# Column Mapping
column_mapping = {
    "Identifier": "ID",
    "Port of Loading": "POL",
    "Port of Discharge": "POD",
    "Container": "CONTAINER",
    "Lumpsum": "FREIGHT",
    "Currency": "Currency",
    "TOTAL SURCHARGE": "Surcharge",
    "Destination": "Destination"
}

st.title("Excel Processing App")

# File uploaders for both received file and template
uploaded_file = st.file_uploader("Upload the received Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload the template Excel file", type=["xlsx"])

if uploaded_file and template_file:
    # Load uploaded file as a DataFrame
    df_uploaded = pd.read_excel(uploaded_file, sheet_name=0)  # Read first sheet

    # Rename columns based on mapping (only apply to existing columns)
    df_uploaded = df_uploaded.rename(columns={k: v for k, v in column_mapping.items() if k in df_uploaded.columns})

    # Ensure "POD" takes "Destination" if available
    if "Destination" in df_uploaded.columns and "POD" in df_uploaded.columns:
        df_uploaded["POD"] = df_uploaded["Destination"].fillna(df_uploaded["POD"])
        df_uploaded.drop(columns=["Destination"], inplace=True)

    # Load the template using openpyxl to preserve formulas
    wb = load_workbook(template_file, data_only=False)  # Keep formulas intact
    ws_feuil1 = wb["Feuil1"]

    # Retrieve Feuil1 headers
    headers = [cell.value for cell in ws_feuil1[1] if cell.value is not None]

    # Match the uploaded data with Feuil1's expected headers
    df_uploaded = df_uploaded.reindex(columns=headers, fill_value="")

    # Clear all rows below the headers in Feuil1
    ws_feuil1.delete_rows(2, ws_feuil1.max_row)

    # Write only the uploaded DataFrame into Feuil1
    for i, row in enumerate(df_uploaded.itertuples(index=False), start=2):
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
