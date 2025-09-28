import pandas as pd

# Define input and output paths
input_file = r"C:\Users\Admin\Desktop\customer_seffi.xlsx"
output_file = "output.xlsx"

# Read the Excel file
df = pd.read_excel(input_file)
print(df.head(30))

