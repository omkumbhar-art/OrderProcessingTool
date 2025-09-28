from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-shefi'  # Change this to a random secret key

# Configuration
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_shefi_excel_file(file_path):
    """
    Process the SHEFI Excel file with the same logic as the original script
    """
    try:
        # Step 1: Read cell A2 from Excel to get PO value
        po_value = pd.read_excel(file_path, header=None, engine='openpyxl').iloc[1, 0]
        
        # Step 2: Read the actual data starting from row 11 (skip first 10 rows)
        df = pd.read_excel(file_path, skiprows=10)
        
        # Select specific columns
        selected_columns = ['VendorStyle#', 'QTY', 'MetalType', 'Color', 'PD#', 'Description', 'Shefi#', 'SHEFIPO#', 'CODE']
        df_selected = df[selected_columns]
        
        # Drop rows with NaN values and make a copy
        df_cleaned = df_selected.dropna().copy()
        
        # Clean newline characters from 'Description'
        df_cleaned['Description'] = df_cleaned['Description'].str.replace('\n', ' ', regex=True)
        
        # Rename columns
        df_cleaned.rename(columns={
            'VendorStyle#': 'StyleCode',
            'QTY': 'OrderQty',
            'MetalType': 'MetalType',
            'Color': 'Tone',
            'PD#': 'ItemRefNo',
            'Description': 'CustomerProductionInstruction',
            'Shefi#': 'SKUNo',
            'SHEFIPO#': 'SHEFIPO#',
            'CODE': 'DIA GRADE'
        }, inplace=True)
        
        # Add SrNo. as the first column
        df_cleaned.insert(loc=0, column='SrNo.', value=range(1, len(df_cleaned) + 1))
        
        # Add ItemSize after StyleCode
        df_cleaned.insert(loc=2, column='ItemSize', value='')
        
        # Insert 'OrderItemPcs' after 'OrderQty' and set all values to 1
        order_qty_index = df_cleaned.columns.get_loc('OrderQty')
        df_cleaned.insert(loc=order_qty_index + 1, column='OrderItemPcs', value=1)
        
        # Add StockType and MakeType after ItemRefNo
        item_ref_index = df_cleaned.columns.get_loc('ItemRefNo')
        df_cleaned.insert(loc=item_ref_index + 1, column='StockType', value='')
        df_cleaned.insert(loc=item_ref_index + 2, column='MakeType', value='')
        
        # Insert 'OrderGroup' and 'Certificate' before 'SKUNo'
        sku_index = df_cleaned.columns.get_loc('SKUNo')
        df_cleaned.insert(loc=sku_index, column='OrderGroup', value='SHEFI')
        df_cleaned.insert(loc=sku_index + 1, column='Certificate', value='')
        
        # Add 8 columns after SKUNo (removed 'null' column from original)
        sku_index = df_cleaned.columns.get_loc('SKUNo')
        new_columns = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for i, col in enumerate(new_columns):
            df_cleaned.insert(loc=sku_index + 1 + i, column=col, value='')
        
        # Insert 'ItemPoNo' after 'Tone'
        tone_index = df_cleaned.columns.get_loc('Tone')
        df_cleaned.insert(loc=tone_index + 1, column='ItemPoNo', value=po_value)
        
        # Create 'Metal' column using: last char of StyleCode + numeric part of MetalType + Tone
        df_cleaned['Metal'] = df_cleaned.apply(
            lambda row: 'G' + re.sub(r'\D', '', str(row['MetalType'])) + str(row['Tone']),
            axis=1
        )
        
        # Replace 'MetalType' column with 'Metal'
        metal_type_index = df_cleaned.columns.get_loc('MetalType')
        df_cleaned.drop(columns=['MetalType'], inplace=True)
        metal_col = df_cleaned.pop('Metal')
        df_cleaned.insert(loc=metal_type_index, column='Metal', value=metal_col)
        
        # Add 'SpecialRemarks' column
        df_cleaned['SpecialRemarks'] = df_cleaned.apply(
            lambda row: f"PD#, {row['ItemRefNo']}, SHEFI # {row['SKUNo']}, SHEFI PO# ,{row['SHEFIPO#']} ,{row['Metal']}, DIA QLTY {row['DIA GRADE']}",
            axis=1
        )
        
        # Insert 'SpecialRemarks' after CustomerProductionInstruction
        dpi_index = df_cleaned.columns.get_loc('CustomerProductionInstruction')
        special_remarks_col = df_cleaned.pop('SpecialRemarks')
        df_cleaned.insert(loc=dpi_index + 1, column='SpecialRemarks', value=special_remarks_col)
        
        # Insert 'DesignProductionInstruction' after 'SpecialRemarks'
        df_cleaned.insert(loc=dpi_index + 2, column='DesignProductionInstruction', value='')
        
        # Define logic for 'StampInstruction'
        def get_stamp_instruction(metal):
            if metal in ["G14W", "G14Y", "G14P", "G14R"]:
                return "14K & DP2 LOGO"
            elif metal in ["G10W", "G10Y", "G10P", "G10R"]:
                return "10K & DP2 LOGO"
            elif metal in ["G18W", "G18Y", "G18P", "G18R"]:
                return "18K & DP2 LOGO"
            elif metal == "PC95":
                return "PT950 & DP2 LOGO"
            elif metal == "A4YUP342-":
                return "ALLOY & DP2 LOGO"
            elif metal == "AG925":
                return "KT & DP2 LOGO"
            else:
                return "0 & DP2 LOGO"
        
        # Insert 'StampInstruction' after 'DesignProductionInstruction'
        df_cleaned['StampInstruction'] = df_cleaned['Metal'].apply(get_stamp_instruction)
        df_cleaned.insert(loc=dpi_index + 3, column='StampInstruction', value=df_cleaned.pop('StampInstruction'))
        
        return df_cleaned, None
        
    except Exception as e:
        return None, str(e)

@app.route('/')
def index():
    return render_template('indexshefi.html')

@app.route('/process', methods=['POST'])
def process_file():
    try:
        # Debug: Print all form data and files
        print("Form data:", dict(request.form))
        print("Files:", dict(request.files))
        
        # Check if file was uploaded
        if 'file' not in request.files:
            print("ERROR: 'file' not in request.files")
            flash('No file part in the request', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        
        print(f"File object: {file}")
        print(f"File filename: '{file.filename}'")
        
        # Validate inputs
        if not file or file.filename == '' or file.filename is None:
            print("ERROR: No file selected or empty filename")
            flash('Please select a file to upload', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            print(f"ERROR: Invalid file type: {file.filename}")
            flash('Invalid file type. Please upload .xlsx or .xls files only', 'error')
            return redirect(url_for('index'))
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        
        # Process the file
        processed_df, error = process_shefi_excel_file(file_path)
        
        if error:
            flash(f'Error processing file: {error}', 'error')
            # Clean up uploaded file
            if os.path.exists(file_path):
                os.remove(file_path)
            return redirect(url_for('index'))
        
        # Save processed file
        output_filename = f"GATI_FORMAT_SHEFI_{timestamp}.xlsx"
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
        processed_df.to_excel(output_path, index=False)
        
        # Clean up uploaded file
        if os.path.exists(file_path):
            os.remove(file_path)
        
        flash('File processed successfully!', 'success')
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        flash(f'An error occurred: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            flash('File not found', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)