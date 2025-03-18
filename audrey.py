import streamlit as st
import pandas as pd
from io import BytesIO

# Define GitHub raw URLs for reference files
PORT_OF_LOADING_URL = "https://raw.githubusercontent.com/kgonciarz/logisticcs/main/reference/port_of_loading_new.xlsx"
PORT_OF_DISCHARGE_URL = "https://raw.githubusercontent.com/kgonciarz/logisticcs/main/reference/port_of_discharge.xlsx"
DETENTION_URL = "https://raw.githubusercontent.com/kgonciarz/logisticcs/main/reference/detention.xlsx"

@st.cache_data(ttl=3600)
def load_reference_files():
    """Fetch reference files from GitHub and cache them."""
    port_of_loading = pd.read_excel(PORT_OF_LOADING_URL)
    port_of_discharge = pd.read_excel(PORT_OF_DISCHARGE_URL)
    detention = pd.read_excel(DETENTION_URL)
    return port_of_loading, port_of_discharge, detention

# Load the reference files once (cached)
port_of_loading, port_of_discharge, detention = load_reference_files()

def process_data(uploaded_file):
    if uploaded_file is None:
        return None
    
    file_rec = pd.read_excel(uploaded_file, sheet_name='Detail', skiprows=17)
    file_rec = file_rec.drop(file_rec.columns[0], axis=1)
    file_rec = file_rec.dropna(subset=['Port of Discharge', 'Destination', 'Port of Loading'], how='all')
    file_rec['Destination'] = file_rec['Destination'].fillna(file_rec['Port of Discharge'])
    
    # Standardize column text formatting
    file_rec['Port of Loading'] = file_rec['Port of Loading'].str.upper()
    file_rec['Destination'] = file_rec['Destination'].str.upper()
    port_of_loading['port_of_loading1'] = port_of_loading['port_of_loading1'].str.upper()
    port_of_discharge['port_of_discharge1'] = port_of_discharge['port_of_discharge1'].str.upper()
    
    # Merge data with reference files
    file_rec = pd.merge(file_rec, port_of_loading, left_on="Port of Loading", right_on="port_of_loading1", how="left")
    file_rec = pd.merge(file_rec, port_of_discharge, left_on="Destination", right_on="port_of_discharge1", how="left")
    
    file_rec['left_container'] = file_rec['Container'].str[:2]
    file_rec = file_rec[~((file_rec['port_of_loading2'] == 'x') | (file_rec['port_of_discharge2'] == 'x'))]
    
    final = file_rec[['port_of_loading2', 'port_of_discharge2', 'left_container']].drop_duplicates()
    lumpsum = file_rec[file_rec['Charge Code'] == 'Lumpsum'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount', 'Curr.']]
    surcharge = file_rec[file_rec['Charge Code'] == 'MFR'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount']]
    ets = file_rec[file_rec['Charge Code'] == 'ETS'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount']]
    
    final = pd.merge(final, lumpsum, on=['port_of_loading2', 'port_of_discharge2', 'left_container'], how="left")
    final.rename(columns={'Amount': 'FREIGHT'}, inplace=True)
    final = pd.merge(final, surcharge, on=['port_of_loading2', 'port_of_discharge2', 'left_container'], how="left")
    final.rename(columns={'Amount': 'Surcharge'}, inplace=True)
    final = pd.merge(final, ets, on=['port_of_loading2', 'port_of_discharge2', 'left_container'], how="left")
    
    final['Surcharge'] = final['Surcharge'].fillna(0) + final['Amount'].fillna(0)
    final = pd.merge(final, detention, left_on=['port_of_discharge2'], right_on=['POD'], how="left")
    
    final['LINER'] = 'not included'
    final['ALL_IN'] = final.apply(
        lambda row: row['FREIGHT'] + row['Surcharge']
        if row['LINER'] == 'not included'
        else row['FREIGHT'] + row['Surcharge'] + row['LINER'], axis=1
    )
    
    final.drop(columns=['POD', 'Amount'], inplace=True)
    
    final.rename(columns={'port_of_loading2': 'POL',
                          'port_of_discharge2': 'POD',
                          'left_container': 'CONTAINER',
                          'Curr.': 'Currency'}, inplace=True)
    
    final = final[['POL', 'POD', 'CONTAINER', 'FREIGHT', 'LINER', 'Currency', 'Surcharge', 'ALL_IN', 'Detention', 'Demurrage']]
    
    return final

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed Data')
    processed_data = output.getvalue()
    return processed_data

st.title("Excel File Processor")

uploaded_file = st.file_uploader("Upload Quotation Excel File", type=["xlsx"])

if st.button("Process File"):
    if uploaded_file:
        final_df = process_data(uploaded_file)
        st.success("File processed successfully!")
        
        st.dataframe(final_df)
        
        excel_data = to_excel(final_df)
        st.download_button(
            label="Download Processed File",
            data=excel_data,
            file_name="Processed_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload the required file.")
