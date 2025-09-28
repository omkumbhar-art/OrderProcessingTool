from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename
import tempfile
from datetime import datetime
import zipfile

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'  # Change this to a random secret key

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

def extract_stamp_text(text):
    """Extract text between 'UFJC 14KY' and '0.70CTW' from StampInstruction"""
    if pd.isna(text):
        return ""
    
    text = str(text)
    # Look for pattern between "UFJC 14KY" and any CTW value (like 0.70CTW)
    pattern = r'UFJC 14KY\s*(.*?)\s*\d+\.\d+CTW'
    match = re.search(pattern, text, re.IGNORECASE)
    
    if match:
        return match.group(1).strip()
    else:
        return ""

def process_excel_file(file_path, po_value, item_no, base_serial_start):
    """
    Process the Excel file with the same logic as the original script
    """
    try:
        # Convert base_serial_start to integer
        base_serial_start = int(base_serial_start)
        
        # Read Excel
        df = pd.read_excel(file_path, skiprows=2)
        selected_columns = ['Serial\nNo', 'Description', 'Stamp', 'Pieces']
        df_selected = df[selected_columns].copy()

        # Clean text
        df_selected['Description'] = df_selected['Description'].str.replace('\n', ' ', regex=True)
        df_selected['Stamp'] = df_selected['Stamp'].str.replace('\n', ' ', regex=True)

        # Rename columns
        df_selected.rename(columns={
            'Serial\nNo': 'SerialNo',
            'Description': 'CustomerProductionInstruction',
            'Pieces': 'OrderItemPcs'
        }, inplace=True)

        # Remove header repeats
        df_selected = df_selected[~df_selected['SerialNo'].isin(['Buyer', 'Serial\nNo'])]

        # Add base columns
        df_selected.insert(0, 'SrNo', range(1, len(df_selected) + 1))
        SrNo_index = df_selected.columns.get_loc('SrNo')
        df_selected.insert(SrNo_index + 1, 'StyleCode', '')
        df_selected.insert(SrNo_index + 2, 'ItemSize', '')
        df_selected.insert(SrNo_index + 3, 'OrderQty', '10')

        # Move OrderItemPcs
        orderqty_index = df_selected.columns.get_loc('OrderQty')
        orderitempcs_col = df_selected.pop('OrderItemPcs')
        df_selected.insert(orderqty_index + 1, 'OrderItemPcs', orderitempcs_col)

        # Metal and Tone
        OrderItemPcs_index = df_selected.columns.get_loc('OrderItemPcs')
        df_selected.insert(OrderItemPcs_index + 1, 'Metal', '')

        def extract_metal(text):
            if pd.notna(text) and '14KY' in text:
                return 'G14Y'
            return ''

        df_selected['Metal'] = df_selected['CustomerProductionInstruction'].apply(extract_metal)
        Metal_index = df_selected.columns.get_loc('Metal')
        df_selected.insert(Metal_index + 1, 'Tone', '')
        df_selected['Tone'] = df_selected['CustomerProductionInstruction'].apply(
            lambda x: 'YG' if pd.notna(x) and '14KY' in x.upper() else ''
        )

        # PO and Ref
        Tone_index = df_selected.columns.get_loc('Tone')
        df_selected.insert(Tone_index + 1, 'ItemPoNo', po_value)
        itempono_index = df_selected.columns.get_loc('ItemPoNo')
        df_selected.insert(itempono_index + 1, 'ItemRefNo', '')
        df_selected.insert(itempono_index + 2, 'StockType', '')
        df_selected.insert(itempono_index + 3, 'MakeType', '')

        # Remarks, StampInstruction
        CustomerProductionInstruction_index = df_selected.columns.get_loc('CustomerProductionInstruction')
        df_selected.insert(CustomerProductionInstruction_index + 1, 'SpecialRemarks', '')
        df_selected.insert(CustomerProductionInstruction_index + 2, 'DesignProductionInstruction', '')
        df_selected.insert(CustomerProductionInstruction_index + 3, 'StampInstruction', '')

        # Order Group and SKU
        Stamp_index = df_selected.columns.get_loc('Stamp')
        df_selected.insert(Stamp_index + 1, 'OrderGroup', '')
        df_selected.insert(Stamp_index + 2, 'Certificate', '')
        df_selected.insert(Stamp_index + 3, 'SKUNo', item_no)

        # Extra fields
        sku_index = df_selected.columns.get_loc('SKUNo')
        new_columns = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for i, col in enumerate(new_columns):
            df_selected.insert(sku_index + 1 + i, col, '')

        # StyleCode
        def generate_style_code(row):
            if pd.notna(row['CustomerProductionInstruction']) and '18IN' in row['CustomerProductionInstruction']:
                tone = row['Tone'] if pd.notna(row['Tone']) else ''
                sku = row['SKUNo'] if pd.notna(row['SKUNo']) else ''
                suffix = 'CO' if 'CO' in sku else ''
                return f"XK2807G-18IN{tone}{suffix}"
            return ''
        df_selected['StyleCode'] = df_selected.apply(generate_style_code, axis=1)

        # SpecialRemarks
        def generate_special_remarks(row):
            remarks = []
            sku = row['SKUNo']
            desc = row['CustomerProductionInstruction']
            if pd.notna(sku): remarks.append(sku)
            if pd.notna(desc) and '14KY' in desc: remarks.append('14K YELLOW GOLD')
            if pd.notna(desc) and '18IN' in desc: remarks.append('SZ 18 INCH')
            remarks.append('DIA QLTY-HI-VS')
            return ','.join(remarks)

        df_selected['SpecialRemarks'] = df_selected.apply(generate_special_remarks, axis=1)

        # StampInstruction group-wise per SrNo
        def generate_stamp_instructions(df, base_serial_start):
            stamp_instructions = []
            for idx, row in df.iterrows():
                srno = row['SrNo']
                start_serial = base_serial_start + (srno - 1) * 10
                end_serial = start_serial + 9

                has_ufjc = 'UFJC' in str(row['Stamp'])
                has_14ky = '14KY' in str(row['CustomerProductionInstruction'])
                qty_is_10 = str(row['OrderQty']) == '10'
                ctw_match = re.search(r'\d+\.\d+CTW', str(row['CustomerProductionInstruction']))
                ctw_value = ctw_match.group() if ctw_match else ''

                if has_ufjc and has_14ky and qty_is_10 and ctw_value:
                    instruction = f"UFJC 14KY {start_serial} to {end_serial} {ctw_value}"
                else:
                    instruction = ''
                stamp_instructions.append(instruction)
            return stamp_instructions

        df_selected['StampInstruction'] = generate_stamp_instructions(df_selected, base_serial_start)

        # Create a dataframe with only ExtractedStamp column
        df_stamp_only = pd.DataFrame()
        df_stamp_only['ExtractedStamp'] = df_selected['StampInstruction'].apply(extract_stamp_text)

        # Drop SerialNo and Stamp columns from original dataframe
        df_selected.drop(columns=['SerialNo'], inplace=True)
        df_selected.drop(columns=['Stamp'], inplace=True)

        return df_selected, df_stamp_only, None

    except Exception as e:
        return None, None, str(e)

@app.route('/')
def index():
    return render_template('indexuneek.html')

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
        po_value = request.form.get('po_value', '').strip()
        item_no = request.form.get('item_no', '').strip()
        base_serial_start = request.form.get('base_serial_start', '').strip()
        
        print(f"File object: {file}")
        print(f"File filename: '{file.filename}'")
        print(f"PO Value: '{po_value}'")
        print(f"Item No: '{item_no}'")
        print(f"Base Serial Start: '{base_serial_start}'")
        
        # Validate inputs
        if not file or file.filename == '' or file.filename is None:
            print("ERROR: No file selected or empty filename")
            flash('Please select a file to upload', 'error')
            return redirect(url_for('index'))
        
        if not po_value or not item_no or not base_serial_start:
            print("ERROR: Missing PO value, Item No, or Base Serial Start")
            flash('Please provide PO value, Item No, and Base Serial Start', 'error')
            return redirect(url_for('index'))
        
        # Validate base_serial_start is numeric
        try:
            int(base_serial_start)
        except ValueError:
            print("ERROR: Base Serial Start is not a valid number")
            flash('Base Serial Start must be a valid number', 'error')
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
        processed_df, processed_df_stamp_only, error = process_excel_file(file_path, po_value, item_no, base_serial_start)
        
        if error:
            flash(f'Error processing file: {error}', 'error')
            # Clean up uploaded file
            if os.path.exists(file_path):
                os.remove(file_path)
            return redirect(url_for('index'))
        
        # Save both processed files
        output_filename_1 = f"GATI_FORMAT_UNEEK_{timestamp}.xlsx"
        output_filename_2 = f"EXTRACTED_STAMP_ONLY_{timestamp}.xlsx"
        zip_filename = f"GATI_FORMAT_UNEEK_FILES_{timestamp}.zip"
        
        output_path_1 = os.path.join(app.config['PROCESSED_FOLDER'], output_filename_1)
        output_path_2 = os.path.join(app.config['PROCESSED_FOLDER'], output_filename_2)
        zip_path = os.path.join(app.config['PROCESSED_FOLDER'], zip_filename)
        
        # Save the Excel files
        processed_df.to_excel(output_path_1, index=False)
        processed_df_stamp_only.to_excel(output_path_2, index=False)
        
        # Create a zip file containing both Excel files
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(output_path_1, output_filename_1)
            zipf.write(output_path_2, output_filename_2)
        
        # Clean up individual Excel files and uploaded file
        if os.path.exists(file_path):
            os.remove(file_path)
        if os.path.exists(output_path_1):
            os.remove(output_path_1)
        if os.path.exists(output_path_2):
            os.remove(output_path_2)
        
        flash('Files processed successfully! Download contains both versions.', 'success')
        return send_file(zip_path, as_attachment=True, download_name=zip_filename)
        
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
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