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
    "TOTAL SURCHARGE": "Surcharge"
}

st.title("Excel Processing App")

uploaded_file = st.file_uploader("Upload the received Excel file", type=["xlsx"])
template_file_path = "/mnt/data/Hapag Q4 Template.xlsx"  # Path to the provided template

if uploaded_file:
    # Load uploaded file and template
    df_uploaded = pd.read_excel(uploaded_file, sheet_name=0)  # Read first sheet
    df_template = pd.read_excel(template_file_path, sheet_name=None)  # Read all sheets
    
    # Rename columns based on mapping
    df_uploaded = df_uploaded.rename(columns=column_mapping)
    
    # Overwrite Feuil1 in the template
    df_template["Feuil1"] = df_uploaded
    
    # Save updated template to a BytesIO buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_template.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    
    # Read Feuil2 after calculations
    df_feuil2 = pd.read_excel(output, sheet_name="Feuil2")
    
    # Display Feuil2
    st.write("### Processed Output (Feuil2)")
    st.dataframe(df_feuil2)
    
    # Provide download option
    st.download_button(
        label="Download Processed File",
        data=output,
        file_name="Processed_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
