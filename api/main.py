from flask import Flask, request, send_file, render_template
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route('/')
def upload_form():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        #These are the uploaded files by the user
        planned_file = request.files['planned']
        actual_file = request.files['actual']
        # pd.read_excel reads the uploaded Excel files into pandas DataFrames
        planned_df = pd.read_excel(planned_file)
        actual_df = pd.read_excel(actual_file)
        
        # Print the columns of both dataframes for inspection
        print("Planned DataFrame columns:", planned_df.columns)
        print("Planned DataFrame first few rows:\n", planned_df.head())
        # pd.read_excel reads the uploaded Excel files into pandas DataFrames
        print("Actual DataFrame columns:", actual_df.columns)
        print("Actual DataFrame first few rows:\n", actual_df.head())
        
        # Adjust column selection based on actual column names
        planned_df = planned_df.iloc[2:, [0, 8]]  # Référence is in A3, Quantité à produire in I3
        actual_df = actual_df.iloc[3:, [13, 19]]  # reference REF1 in N4, Pcs produites in T4
        
        # Rename the columns for consistency
        planned_df.columns = ['Product', 'Planned Quantity']
        actual_df.columns = ['Product', 'Actual Quantity']
        
        # Trim whitespace and convert product references to uppercase to ensure consistency
        planned_df['Product'] = planned_df['Product'].str.upper().str.strip()
        actual_df['Product'] = actual_df['Product'].str.upper().str.strip()
        
        # Convert quantities to numeric, handling any potential errors
        planned_df['Planned Quantity'] = pd.to_numeric(planned_df['Planned Quantity'], errors='coerce')
        actual_df['Actual Quantity'] = pd.to_numeric(actual_df['Actual Quantity'], errors='coerce')
        
        # Drop rows with NaN values in 'Planned Quantity' or 'Actual Quantity'
        planned_df = planned_df.dropna(subset=['Planned Quantity'])
        actual_df = actual_df.dropna(subset=['Actual Quantity'])
        
        # Sum actual quantities by product to get total actual production
        actual_df = actual_df.groupby('Product', as_index=False)['Actual Quantity'].sum()
        
        # Merge the dataframes on the 'Product' column
        merged_df = pd.merge(planned_df, actual_df, on='Product', how='outer')
        
        # Fill NaN values with 0 for 'Planned Quantity' and 'Actual Quantity'
        merged_df['Planned Quantity'] = merged_df['Planned Quantity'].fillna(0)
        merged_df['Actual Quantity'] = merged_df['Actual Quantity'].fillna(0)
        
        # Calculate the adherence rate
        merged_df['Adherence Rate'] = (merged_df['Actual Quantity'] / merged_df['Planned Quantity']) * 100
        
        # Replace infinite values in 'Adherence Rate' with NaN
        merged_df['Adherence Rate'].replace([float('inf'), -float('inf')], pd.NA, inplace=True)
        
        # Drop rows with NaN values in 'Adherence Rate'
        merged_df = merged_df.dropna(subset=['Adherence Rate'])
        
        # Print the merged dataframe for debugging
        print("Merged DataFrame:")
        print(merged_df)
        
        # Save the results to a new Excel file in a temporary directory
        temp_dir = tempfile.mkdtemp()
        output_path = os.path.join(temp_dir, 'Adherence_Rate.xlsx')
        merged_df.to_excel(output_path, index=False)
        
        return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
