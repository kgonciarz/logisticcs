import pandas as pd
import os
import shutil
import time

# --- Step 1: Define File Paths ---
template_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\Hapag Q4 Template.xlsx"
output_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\Processed_Feuil2.xlsx"
uploaded_file_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\Quotation_Q2408BSL00081_HAPAGL_066.xlsx"
aaa_file = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\AAA_freight_test.xlsx"

print("Appending data to AAA_freight_test.xlsx...")

# Load the processed file
df_processed = pd.read_excel(output_path)

# Load the AAA freight test file
df_aaa = pd.read_excel(aaa_file)

# ✅ Convert column names to strings before stripping spaces
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

# ✅ Ensure only available columns are selected to avoid KeyError
valid_columns = [col for col in column_mapping.keys() if col in df_processed.columns]
df_to_append = df_processed[valid_columns].rename(columns={k: column_mapping[k] for k in valid_columns})

# Ensure "FREIGHT" column has a consistent numeric format
if "FREIGHT" in df_aaa.columns:
    df_aaa["FREIGHT"] = pd.to_numeric(df_aaa["FREIGHT"], errors="coerce")  # Convert existing FREIGHT to numbers

if "FREIGHT" in df_to_append.columns:
    df_to_append["FREIGHT"] = pd.to_numeric(df_to_append["FREIGHT"], errors="coerce")  # Convert new FREIGHT data to numbers

# Append the processed data to AAA freight test without overwriting existing data
df_aaa = pd.concat([df_aaa, df_to_append], ignore_index=True)

# Overwrite the original AAA_freight_test.xlsx with updated data
df_aaa.to_excel(aaa_file, index=False)

print("✅ Data successfully appended to AAA_freight_test.xlsx with FREIGHT as numeric.")
