import pandas as pd

file_path = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\Quotation_Q2408BSL00081_HAPAGL_066.xlsx"
file_rec = pd.read_excel(file_path, sheet_name='Detail', skiprows=17)
file_rec = file_rec.drop(file_rec.columns[0], axis=1)

file_rec = file_rec.dropna(subset=['Port of Discharge', 'Destination', 'Port of Loading'], how='all')

file_rec['Destination'] = file_rec['Destination'].fillna(file_rec['Port of Discharge'])

file_path1 = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\reference\port_of_loading.xlsx"
port_of_loading = pd.read_excel(file_path1)

file_path2 = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\reference\port_of_discharge.xlsx"
port_of_discharge = pd.read_excel(file_path2)

file_rec['Port of Loading'] = file_rec['Port of Loading'].str.upper()
file_rec['Destination'] = file_rec['Destination'].str.upper()
port_of_loading['port_of_loading1'] = port_of_loading['port_of_loading1'].str.upper()
port_of_discharge['port_of_discharge1'] = port_of_discharge['port_of_discharge1'].str.upper()

file_rec = pd.merge(file_rec, port_of_loading, left_on="Port of Loading", right_on="port_of_loading1", how="left")
file_rec = pd.merge(file_rec, port_of_discharge, left_on="Destination", right_on="port_of_discharge1", how="left")

file_rec['left_container'] = file_rec['Container'].str[:2]

file_rec= file_rec[~((file_rec['port_of_loading2'] == 'x') | (file_rec['port_of_discharge2'] == 'x'))]

final = file_rec[['port_of_loading2', 'port_of_discharge2', 'left_container']].drop_duplicates()
lumpsum = file_rec[file_rec['Charge Code'] == 'Lumpsum'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount', 'Curr.']]
surcharge = file_rec[file_rec['Charge Code'] == 'MFR'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount']]
ets = file_rec[file_rec['Charge Code'] == 'ETS'][['port_of_loading2', 'port_of_discharge2', 'left_container', 'Amount']]

final = pd.merge(final, lumpsum, left_on=['port_of_loading2', 'port_of_discharge2', 'left_container'], right_on=['port_of_loading2', 'port_of_discharge2', 'left_container'], how="left")
final.rename(columns={'Amount': 'FREIGHT'}, inplace=True)
final = pd.merge(final, surcharge, left_on=['port_of_loading2', 'port_of_discharge2', 'left_container'], right_on=['port_of_loading2', 'port_of_discharge2', 'left_container'], how="left")
final.rename(columns={'Amount': 'Surcharge'}, inplace=True)
final = pd.merge(final, ets, left_on=['port_of_loading2', 'port_of_discharge2', 'left_container'], right_on=['port_of_loading2', 'port_of_discharge2', 'left_container'],how="left")

final['Surcharge'] = final['Surcharge'].fillna(0) + final['Amount'].fillna(0)

file_path3 = r"C:\Users\Klaudia Gonciarz\OneDrive - Cocoasource SA\Documents\Audrey\reference\detention.xlsx"
detention = pd.read_excel(file_path3)
final = pd.merge(final, detention, left_on=['port_of_discharge2'], right_on=['POD'],how="left")

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




