import pandas as pd
import glob

# Set the folder containing the reports and output file name
input_folder = "./input_reports/"  # Replace with the folder path
output_file = 'consolidated_report.xlsx'

# Get the list of Excel files in the folder
excel_files = sorted(glob.glob(input_folder + "*.xlsx"))  # Ensure files are sorted by name

# Initialize a DataFrame for Report 1 to use as the base order
base_df = pd.DataFrame()

# Initialize an empty DataFrame to store the merged data
merged_df = pd.DataFrame()

# Loop through each file and merge data
for i, file in enumerate(excel_files, start=1):
    # Read the file into a DataFrame and rename columns explicitly
    df = pd.read_excel(file, header=0)  # Assuming the first row has headers
    df.columns = ['Transactions', f'Report {i} time(90%)']  # Force consistent column names
    
    if i == 1:
        # Use the first report as the base for order
        base_df = df
    
    # Merge with the main DataFrame
    if merged_df.empty:
        merged_df = df
    else:
        merged_df = pd.merge(merged_df, df, on='Transactions', how='outer')

# Sort by the base DataFrame's order
merged_df = base_df[['Transactions']].merge(merged_df, on='Transactions', how='left')

# Replace NaN with empty strings if needed
merged_df.fillna("", inplace=True)

# Reorder columns so the reports are in order (Report 1, Report 2, ..., Report 5)
columns_order = ['Transactions'] + [f'Report {i} time(90%)' for i in range(1, len(excel_files) + 1)]
merged_df = merged_df[columns_order]

# Write the merged DataFrame to an Excel file
merged_df.to_excel(output_file, index=False)

print(f"Consolidated report saved to {output_file}")

