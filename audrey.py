import streamlit as st
import pandas as pd
import io

# Column Mapping
column_mapping = {
    "Identifier": "ID",
    "Port of Loading": "POL",
    "Port of Discharge": "POD",
    "Container": "CONTAINER",
    "Lumpsum": "FREIGHT",
    "Currency": "Currency",
    "TOTAL SURCHARGE": "Surcharge",
    "Destination": "Destination"  # Ensure Destination is included
}

st.title("Excel Processing App")

# File uploaders for both received file and template
uploaded_file = st.file_uploader("Upload the received Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload the template Excel file", type=["xlsx"])

if uploaded_file and template_file:
    # Load uploaded files
    df_uploaded = pd.read_excel(uploaded_file, sheet_name=0)  # Read first sheet
    df_template = pd.read_excel(template_file, sheet_name=None, engine='openpyxl', keep_default_na=False)  # Read all sheets including formulas

    # Rename columns based on mapping
    df_uploaded = df_uploaded.rename(columns=column_mapping)

    # Ensure columns exist before processing
    if "Destination" in df_uploaded.columns and "POD" in df_uploaded.columns:
        df_uploaded["POD"] = df_uploaded["Destination"].fillna(df_uploaded["POD"])
        df_uploaded.drop(columns=["Destination"], inplace=True)

    # Overwrite Feuil1 in the template
    df_template["Feuil1"] = df_uploaded

    # Save updated template to a BytesIO buffer using openpyxl to preserve formulas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_template.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    # Provide download option with formulas preserved
    st.download_button(
        label="Download Processed File with Formulas",
        data=output,
        file_name="Processed_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
