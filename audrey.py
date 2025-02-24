import pandas as pd
import streamlit as st
import os

# --- Step 1: Define File Paths ---
processed_file_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\Processed_Feuil2.xlsx"
aaa_file_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\AAA_freight_test.xlsx"

# Streamlit UI
st.title("📤 AAA Freight Data Processing App")
st.write("This app appends processed data from `Processed_Feuil2.xlsx` to `AAA_freight_test.xlsx`.")

# === ✅ Step 1: Load Files ===
if os.path.exists(processed_file_path) and os.path.exists(aaa_file_path):

    st.success("✅ Required files found!")

    # Load the processed file
    df_processed = pd.read_excel(processed_file_path)

    # Load the AAA freight test file
    df_aaa = pd.read_excel(aaa_file_path)

    # ✅ Convert column names to strings before stripping spaces
    df_aaa.columns = df_aaa.columns.astype(str).str.strip()
    df_processed.columns = df_processed.columns.astype(str).str.strip()

    # ✅ Mapping of columns from `Processed_Feuil2.xlsx` to `AAA_freight_test.xlsx`
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

    # ✅ Append the processed data to AAA freight test without overwriting existing data
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
else:
    st.error("⚠️ One or more required files are missing! Please check file paths.")
