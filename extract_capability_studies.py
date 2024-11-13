import openpyxl
import pandas as pd

excel_file = 'P-V12 - Capabilité cote 1.6 ±0.03 MIN-HAUT.xlsm'
workbook = openpyxl.load_workbook(excel_file, data_only=True)
#

sheet = workbook['Synthesis']

reference = sheet['G4'].value   
dimension = sheet['G5'].value  
cp = sheet['B22'].value 
cpk = sheet['B23'].value 

data = {
    'Reference': [reference],
    'Dimension': [dimension],
    'Cp': [cp],
    'Cpk': [cpk]
}
table_df = pd.DataFrame(data)

print(table_df)


table_df.to_excel('extracted_data.xlsx', index=False)
