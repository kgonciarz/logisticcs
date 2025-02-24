import pandas as pd
import streamlit as st
import os

# --- Define File Paths for the AAA Freight Test ---
aaa_file_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\AAA_freight_test.xlsx"

# --- Streamlit UI ---
st.title("📤 Upload & Process Your Freight Data")
st.write("Upload your Excel file, and the system will process and append it to AAA Freight Test.")

# === ✅ Step 1: File Upload ===
uploaded_file = st.file_uploader("Upload Your Processed File (Excel)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ File uploaded successfully! Processing...")

    # Load the uploaded file into a Pandas DataFrame
    df_processed = pd.read_excel(uploaded_file)

    # Ensure the AAA Freight Test file exists before processing
    if os.path.exists(aaa_file_path):
        df_aaa = pd.read_excel(aaa_file_path)
    else:
        st.error("⚠️ The AAA Freight Test file does not exist! Please check the file path.")
        st.stop()

    # ✅ Convert column names to strings before stripping spaces
    df_aaa.columns = df_aaa.columns.astype(str).str.strip()
    df_processed.columns = df_processed.columns.astype(str).str.strip()

    # ✅ Mapping of columns from the uploaded file to AAA_freight_test.xlsx
    column_mapping = {
        "Identifier": "ID",
        "Port of Loading": "POL",
        "Port of Discharge": "POD",
        "Container": "CONTAINER",
        "Lumpsum": "FREIGHT",  # Ensure correct data type
        "Currency": "Currency",
        "TOTAL SURCHARGE": "Surcharge"
    }

    # ✅ Ensure only available columns are selected to avoid KeyError
    valid_columns = [col for col in column_mapping.keys() if col in df_processed.columns]
    df_to_append = df_processed[valid_columns].rename(columns={k: column_mapping[k] for k in valid_columns})

    # ✅ Ensure "FREIGHT" column has a consistent numeric format
    if "FREIGHT" in df_aaa.columns:
        df_aaa["FREIGHT"] = pd.to_numeric(df_aaa["FREIGHT"], errors="coerce")  # Convert existing FREIGHT to numbers

    if "FREIGHT" in df_to_append.columns:
        df_to_append["FREIGHT"] = pd.to_numeric(df_to_append["FREIGHT"], errors="coerce")  # Convert new FREIGHT data to numbers

    # ✅ Append the processed data to AAA Freight Test without overwriting existing data
    df_aaa = pd.concat([df_aaa, df_to_append], ignore_index=True)

    # ✅ Save the updated AAA Freight Test file
    df_aaa.to_excel(aaa_file_path, index=False)

    st.success("✅ Data successfully appended to AAA Freight Test!")

    # ✅ Provide a Download Button for the Updated File
    with open(aaa_file_path, "rb") as f:
        st.download_button(
            label="📥 Download Updated AAA Freight Test",
            data=f,
            file_name="AAA_freight_test.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
