import pandas as pd
import os

# Input folder and output folder
input_folder = "./input_reports"  # Folder where your Excel files are located
output_file = "./output_reports/Consolidated_Filtered_Reports.csv"  # Single output CSV file

# Ensure output folder exists
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# Initialize a list to store all rows for the consolidated CSV
all_data = []

# Process each Excel file in the input folder
report_number = 1
for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx"):
        # Read the Excel file
        file_path = os.path.join(input_folder, file_name)
        df = pd.read_excel(file_path)

        # Ensure the "time(90%)" column is numeric
        df["time(90%)"] = pd.to_numeric(df["time(90%)"], errors="coerce")

        # Drop rows with invalid "time(90%)" values
        df = df.dropna(subset=["time(90%)"])

        # Filter transactions: 1.8 to <2 seconds
        filtered_1_8_to_2 = df[(df["time(90%)"] >= 1.8) & (df["time(90%)"] < 2)].copy()

        # Filter transactions: >=2 seconds
        filtered_2_or_more = df[df["time(90%)"] >= 2].copy()

        # Add a heading for the current report
        all_data.append([f"Report {report_number}: Filtered Data for {file_name}"])
        
        # Add the first filtered table for 1.8 to <2 seconds
        all_data.append(["Table: Transactions with 1.8 to <2 seconds"])
        all_data.extend(filtered_1_8_to_2.values.tolist())
        all_data.append([])  # Add a blank row for separation
        
        # Add the second filtered table for >=2 seconds
        all_data.append(["Table: Transactions with >=2 seconds"])
        all_data.extend(filtered_2_or_more.values.tolist())
        all_data.append([])  # Add another blank row for separation

        report_number += 1

# Save the consolidated data to a CSV file
with open(output_file, "w", encoding="utf-8") as f:
    for row in all_data:
        f.write(",".join(map(str, row)) + "\n")

print(f"Processing complete. Consolidated report saved as {output_file}.")
