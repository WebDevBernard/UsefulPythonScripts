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


# Function to compare two DataFrames
def compare_dataframes(df1, df2, output_file):
    # Merge the two DataFrames on 'policynum'
    merged = pd.merge(df1, df2, on='policynum', suffixes=('_df1', '_df2'))

    # Filter rows where ccode differs
    filtered = merged[merged['ccode_df1'] != merged['ccode_df2']]

    # Save the result if there are differences
    if not filtered.empty:
        filtered.to_excel(output_file, index=False)
        print(f"Differences saved to: {output_file}")


# Iterate through pairs of Excel files
for i, file1 in enumerate(files):
    for file2 in files[i + 1:]:
        try:
            # Read the Excel files
            df1 = (
                pd.read_excel(file1, engine="xlrd")
                if file1.suffix.lower() == ".xls"
                else pd.read_excel(file1, engine="openpyxl")
            )
            df2 = (
                pd.read_excel(file2, engine="xlrd")
                if file2.suffix.lower() == ".xls"
                else pd.read_excel(file2, engine="openpyxl")
            )

            # Ensure required columns exist in both files
            if {'policynum', 'ccode'}.issubset(df1.columns) and {'policynum', 'ccode'}.issubset(df2.columns):
                # Define output file path
                output_file = output_dir / f"Comparison_{file1.stem}_vs_{file2.stem}.xlsx"

                # Compare the DataFrames
                compare_dataframes(df1, df2, output_file)

            else:
                print(f"Skipping files {file1.name} and {file2.name}: Missing required columns")
        except Exception as e:
            print(f"Error processing files {file1.name} and {file2.name}: {e}")
