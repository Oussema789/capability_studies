import os
import pandas as pd
import openpyxl

root_directory = r"C:/Users/Oussema KHELIFI/OneDrive/Desktop/Capabilité"

if not os.path.exists(root_directory):
    print(f"Root directory '{root_directory}' not found.")
else:
    print(f"Root directory '{root_directory}' found.")

summary_data = []


for emp_folder in ['EMP 1', 'EMP 2', 'EMP 3', 'EMP 4']:
    emp_path = os.path.join(root_directory, emp_folder)
    
    
    if not os.path.exists(emp_path):
        print(f"Folder {emp_folder} not found at path: {emp_path}")
        continue
    else:
        print(f"Processing folder: {emp_folder}")

    
    for file_name in os.listdir(emp_path):
        if file_name.endswith('.xlsm') or file_name.endswith('.xlsx'):
            file_path = os.path.join(emp_path, file_name)
            print(f"Processing file: {file_path}")
            
           
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            sheet = workbook['Synthesis']  
            
            
            reference = sheet['G4'].value  
            dimension = sheet['G5'].value  
            cp = sheet['B22'].value  
            cpk = sheet['B23'].value  

       
            summary_data.append({
                'EMP Folder': emp_folder,
                'File Name': file_name,
                'Reference': reference,
                'Dimension': dimension,
                'Cp': cp,
                'Cpk': cpk
            })


summary_df = pd.DataFrame(summary_data)


desktop_path = r"C:/Users/Oussema KHELIFI/OneDrive/Desktop/EMP_summary.xlsx"
summary_df.to_excel(desktop_path, index=False)

print(f"Summary saved to {desktop_path}")
