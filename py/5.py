import os
import pandas as pd
from pathlib import Path

# Define input directory
input_dir = Path("E:\\PythonProjects\\UsefulPythonScripts\\input")

# Get list of all Excel files (.xlsx and .xls)
xlsx_files = input_dir.glob("*.xlsx")
xls_files = input_dir.glob("*.xls")
files = list(xlsx_files) + list(xls_files)

# Define output directory (Desktop)
output_dir = Path.home() / "Desktop"
output_dir.mkdir(parents=True, exist_ok=True)

# Initialize an empty DataFrame to store the combined results
combined_df = pd.DataFrame()

# Loop over each file in the list
for i, file1 in enumerate(files):
    # Read the Excel file
    df = (
        pd.read_excel(file1, engine="xlrd")  # For .xls files
        if file1.suffix.lower() == ".xls"
        else pd.read_excel(file1, engine="openpyxl")  # For .xlsx files
    )

    # Drop duplicates based on the 'policynum' column and keep the first occurrence
    df = df.drop_duplicates(subset=["policynum"], keep='first')

    # Append the current DataFrame to the combined DataFrame
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# Define the output file path
output_file = output_dir / "test.xlsx"

# Save the combined DataFrame to an Excel file
combined_df.to_excel(output_file, index=False)

print(f"Processed file saved to {output_file}")
