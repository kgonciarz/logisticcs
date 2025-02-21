import streamlit as st
import pandas as pd
import os

# --- Step 1: Streamlit UI ---
st.title("AAA Freight Processor 🚢")
st.write("Upload the necessary files and process them.")

# File uploaders for user inputs
uploaded_template = st.file_uploader("Upload Hapag Q4 Template", type=["xlsx"])
uploaded_quotation = st.file_uploader("Upload Quotation File", type=["xlsx"])
uploaded_aaa = st.file_uploader("Upload AAA Freight Test File", type=["xlsx"])

# Button to trigger processing
if st.button("Process Files"):
    if uploaded_template and uploaded_quotation and uploaded_aaa:
        # --- Step 2: Read Uploaded Files ---
df_template = pd.ExcelFile(uploaded_template, engine="openpyxl")
df_quotation = pd.ExcelFile(uploaded_quotation, engine="openpyxl")
        df_aaa = pd.read_excel(uploaded_aaa, engine="openpyxl")

        # Load relevant sheets
        df_processed = pd.read_excel(df_template, sheet_name="Feuil2")
        df_uploaded = pd.read_excel(df_quotation, sheet_name="Detail")

        # --- Step 3: Process Data ---
        df_aaa.columns = df_aaa.columns.astype(str).str.strip()
        df_processed.columns = df_processed.columns.astype(str).str.strip()

        # Mapping of columns from Processed_Feuil2 to AAA_freight_test
        column_mapping = {
            "Identifier": "ID",
            "Port of Loading": "POL",
            "Port of Discharge": "POD",
            "Container": "CONTAINER",
            "Lumpsum": "FREIGHT",  # Ensure correct data type
            "Currency": "Currency",
            "TOTAL SURCHARGE": "Surcharge"
        }

        # Ensure only available columns are selected
        valid_columns = [col for col in column_mapping.keys() if col in df_processed.columns]
        df_to_append = df_processed[valid_columns].rename(columns={k: column_mapping[k] for k in valid_columns})

        # Convert FREIGHT column to numeric
        if "FREIGHT" in df_aaa.columns:
            df_aaa["FREIGHT"] = pd.to_numeric(df_aaa["FREIGHT"], errors="coerce")

        if "FREIGHT" in df_to_append.columns:
            df_to_append["FREIGHT"] = pd.to_numeric(df_to_append["FREIGHT"], errors="coerce")

        # Append the processed data to AAA freight test
        df_aaa = pd.concat([df_aaa, df_to_append], ignore_index=True)

        # --- Step 4: Save Processed File ---
        output_path = "Processed_AAA_Freight_Test.xlsx"
        df_aaa.to_excel(output_path, index=False)

        # Display success message and show processed data
        st.success("✅ Data successfully appended to AAA_freight_test.xlsx with FREIGHT as numeric.")
        st.write(df_aaa.head())

        # Provide download button
        with open(output_path, "rb") as file:
            st.download_button(
                label="Download Processed File 📂",
                data=file,
                file_name="Processed_AAA_Freight_Test.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("❌ Please upload all required files.")
