import pandas as pd
import streamlit as st

def extract_distinct_combinations(file_path, sheet_name):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    
    # Load the specified sheet
    df_raw = xls.parse(sheet_name=sheet_name)
    
    # Identify the correct header row dynamically
    for i, row in df_raw.iterrows():
        if row.notna().sum() > 5:  # Assuming at least 5 non-null values indicate a header row
            header_row = i
            break
    
    # Reload data using the identified header row
    df_cleaned = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=header_row)
    
    # Rename columns using the first row as headers and drop the original header row
    df_cleaned.columns = df_cleaned.iloc[0]
    df_cleaned = df_cleaned[1:].reset_index(drop=True)
    
    # Ensure only valid rows are used where all required columns are present in the same row
    df_cleaned = df_cleaned.dropna(subset=["Port of Loading", "Port of Discharge", "Container"])
    
    # Selecting the relevant columns from the same row
    df_selected = df_cleaned[["Port of Loading", "Port of Discharge", "Container"]]
    
    # Dropping duplicates to get distinct combinations strictly from the same rows
    df_distinct = df_selected.drop_duplicates().reset_index(drop=True)
    
    return df_distinct

# Streamlit App
st.title("Extract Distinct Port and Container Combinations")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    sheet_name = "Detail"  # Ensure correct sheet name
    df_distinct = extract_distinct_combinations(uploaded_file, sheet_name)
    
    st.write("### Distinct Combinations")
    st.dataframe(df_distinct)
    
    # Provide download option
    df_distinct.to_excel("distinct_combinations.xlsx", index=False)
    st.download_button(
        label="Download Excel File",
        data=open("distinct_combinations.xlsx", "rb").read(),
        file_name="distinct_combinations.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
