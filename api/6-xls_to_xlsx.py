from helper_funtions import emoji
from pathlib import Path
import pandas as pd

base_dir = Path(Path(__file__).parent.parent / "input")
input_dir = base_dir / "xls to xlsx"
xls_files = input_dir.glob("*.xls")
xls_file_paths = []

for xls_file in xls_files:
    file_path = str(xls_file)
    xls_file_paths.append(file_path)

for xls_file in xls_file_paths:
    print(f"\n{emoji}   XLSX_FILENAME: {xls_file} {emoji}\n")
    input_path = Path(xls_file)
    output_path = base_dir / f"{Path(xls_file).stem}.xlsx"
    df = pd.read_excel(input_path, engine="xlrd")
    writer = pd.ExcelWriter(output_path, engine="openpyxl")
    df.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.close()
