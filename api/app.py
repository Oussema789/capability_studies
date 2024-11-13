from flask import Flask, request, send_file, render_template, flash
import os
import pandas as pd
import openpyxl
import tempfile
import zipfile

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Needed for flashing messages

# Constants for folder names and expected sheet names
EMP_FOLDERS = ['EMP 1', 'EMP 2', 'EMP 3', 'EMP 4']
EXPECTED_SHEET_NAME = 'Synthesis'
OUTPUT_FILE_NAME = 'EMP_summary.xlsx'
MAX_FILE_SIZE_MB = 4  # Maximum file size in MB

@app.route('/')
def upload_form():
    return render_template('upload_form.html')  # Render form to upload ZIP file

def validate_zip_file(zip_file):
    """Validate the uploaded ZIP file."""
    if not zip_file:
        flash("No file part in the request.")
        return False
    if zip_file.filename == '':
        flash("No selected file.")
        return False
    if not zip_file.filename.endswith('.zip'):
        flash("Please upload a ZIP file.")
        return False
    return True

def extract_zip(zip_path, temp_dir):
    """Extract the ZIP file to a temporary directory."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

def process_excel_file(file_path):
    """Process an Excel file and extract relevant data."""
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    
    if EXPECTED_SHEET_NAME not in workbook.sheetnames:
        flash(f"'{EXPECTED_SHEET_NAME}' sheet not found in {file_path}")
        return None

    sheet = workbook[EXPECTED_SHEET_NAME]
    
    # Retrieve values from specified cells
    reference = sheet['G4'].value
    dimension = sheet['G5'].value
    cp = sheet['B22'].value
    cpk = sheet['B23'].value
    
    return {
        'Reference': reference,
        'Dimension': dimension,
        'Cp': cp,
        'Cpk': cpk
    }

@app.route('/process', methods=['POST'])
def process_zip():
    zip_file = request.files.get('zip_file')

    # Validate the uploaded ZIP file
    if not validate_zip_file(zip_file):
        return render_template('upload_form.html')

    # Check if the uploaded ZIP file size exceeds the limit
    if zip_file.content_length > MAX_FILE_SIZE_MB * 1024 * 1024:
        flash(f"The uploaded ZIP file exceeds the maximum size of {MAX_FILE_SIZE_MB} MB.")
        return render_template('upload_form.html')

    # Create a temporary directory to extract the ZIP content
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, zip_file.filename)
    zip_file.save(zip_path)

    # Extract the ZIP file contents
    extract_zip(zip_path, temp_dir)

    # Identify the inner folder based on the ZIP filename
    inner_folder_name = os.path.splitext(zip_file.filename)[0]
    inner_folder_path = os.path.join(temp_dir, inner_folder_name)

    if not os.path.exists(inner_folder_path):
        flash(f"Expected folder '{inner_folder_name}' not found in ZIP file.")
        return render_template('upload_form.html')

    summary_data = []

    # Process each EMP folder inside the inner extracted folder
    for emp_folder in EMP_FOLDERS:
        emp_path = os.path.join(inner_folder_path, emp_folder)
        
        if not os.path.exists(emp_path):
            flash(f"Folder '{emp_folder}' not found in ZIP file.")
            continue

        for file_name in os.listdir(emp_path):
            if file_name.endswith(('.xlsm', '.xlsx')):
                file_path = os.path.join(emp_path, file_name)

                # Check if individual Excel file size exceeds limit before processing
                if os.path.getsize(file_path) > MAX_FILE_SIZE_MB * 1024 * 1024:
                    flash(f"The file '{file_name}' exceeds the maximum size of {MAX_FILE_SIZE_MB} MB and will be skipped.")
                    continue
                
                data = process_excel_file(file_path)
                
                if data is None or all(value is None for value in data.values()):
                    flash(f"No data found in {file_name} in folder {emp_folder}.")
                else:
                    summary_data.append({
                        'EMP Folder': emp_folder,
                        'File Name': file_name,
                        **data  # Unpack dictionary to include values directly
                    })

    # Create a DataFrame and save to an Excel file
    summary_df = pd.DataFrame(summary_data)
    
    output_path = os.path.join(temp_dir, OUTPUT_FILE_NAME)
    
    summary_df.to_excel(output_path, index=False)

    # Provide feedback based on summary data availability
    if summary_df.empty:
        flash("No data was found in any of the files.")
    else:
        flash("Summary file created successfully.")

        # Check output size before sending it back to user
        if os.path.getsize(output_path) > MAX_FILE_SIZE_MB * 1024 * 1024:
            flash(f"The generated summary exceeds {MAX_FILE_SIZE_MB} MB. Please reduce input sizes.")
            return render_template('upload_form.html')

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)