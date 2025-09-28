import pandas as pd
import os

def extract_vendorstyle_column(file_path, output_file=None):
    """
    Extract VendorStyle# column from Excel file
    
    Args:
        file_path (str): Path to the Excel file
        output_file (str, optional): Path to save extracted data as CSV
    
    Returns:
        pandas.Series: The VendorStyle# column data
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Print column names to help identify the exact column name
        print("Available columns:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        
        # Look for VendorStyle# column (case-insensitive and flexible matching)
        vendorstyle_col = None
        for col in df.columns:
            if 'vendorstyle' in col.lower() and '#' in col:
                vendorstyle_col = col
                break
        
        if vendorstyle_col is None:
            print("VendorStyle# column not found!")
            print("Please check the column names above and modify the code accordingly.")
            return None
        
        # Extract the column
        extracted_data = df[vendorstyle_col]
        
        print(f"\nExtracted {len(extracted_data)} rows from column '{vendorstyle_col}'")
        print(f"Non-null values: {extracted_data.notna().sum()}")
        
        # Display first few values
        print(f"\nFirst 10 values:")
        print(extracted_data.head(10))
        
        # Save to CSV if output file specified
        if output_file:
            extracted_data.to_csv(output_file, index=False, header=True)
            print(f"\nData saved to: {output_file}")
        
        return extracted_data
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error reading file: {str(e)}")
        return None

def extract_multiple_columns(file_path, column_patterns, output_file=None):
    """
    Extract multiple columns that match certain patterns
    
    Args:
        file_path (str): Path to the Excel file
        column_patterns (list): List of column name patterns to match
        output_file (str, optional): Path to save extracted data as CSV
    
    Returns:
        pandas.DataFrame: DataFrame with matching columns
    """
    try:
        df = pd.read_excel(file_path)
        
        # Find columns that match any of the patterns
        matching_cols = []
        for col in df.columns:
            for pattern in column_patterns:
                if pattern.lower() in col.lower():
                    matching_cols.append(col)
                    break
        
        if not matching_cols:
            print("No matching columns found!")
            return None
        
        # Extract matching columns
        extracted_df = df[matching_cols]
        
        print(f"Extracted columns: {matching_cols}")
        print(f"Shape: {extracted_df.shape}")
        
        # Save to CSV if output file specified
        if output_file:
            extracted_df.to_csv(output_file, index=False)
            print(f"Data saved to: {output_file}")
        
        return extracted_df
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

# Example usage
if __name__ == "__main__":
    # Replace with your Excel file path
    excel_file = "your_file.xlsx"
    
    # Method 1: Extract just VendorStyle# column
    print("=== Extracting VendorStyle# Column ===")
    vendorstyle_data = extract_vendorstyle_column(
        file_path=excel_file,
        output_file="vendorstyle_column.csv"
    )
    
    # Method 2: Extract multiple related columns
    print("\n=== Extracting Multiple Columns ===")
    patterns = ["vendorstyle", "vendor", "style"]  # Add more patterns as needed
    multiple_cols = extract_multiple_columns(
        file_path=excel_file,
        column_patterns=patterns,
        output_file="vendor_related_columns.csv"
    )
    
    # Method 3: Direct pandas approach (if you know exact column name)
    print("\n=== Direct Extraction (if you know exact column name) ===")
    try:
        df = pd.read_excel(excel_file)
        
        # Replace 'VendorStyle#' with the exact column name from your file
        if 'VendorStyle#' in df.columns:
            vendor_style_col = df['VendorStyle#']
            print(f"Extracted {len(vendor_style_col)} values")
            print(vendor_style_col.head())
            
            # Save to file
            vendor_style_col.to_csv("direct_extraction.csv", index=False, header=True)
        else:
            print("Column 'VendorStyle#' not found. Available columns:")
            print(list(df.columns))
            
    except Exception as e:
        print(f"Error in direct extraction: {str(e)}")