from helper_functions import base_dir, emoji, unique_file_name
from pdfplumber_functions import plumber_extract_tables
import pandas as pd
from pathlib import Path
import os
# Updates renewal List from EDI LOG work in progress, EDILOG needs to be placed in input folder to work

# Table Strategy for identifying table, using h-lines or h-text
ts = {
    "vertical_strategy": "explicit",
    "explicit_vertical_lines": [23.64, 50.64, 106.91, 159.36, 217.94, 341.64, 392.67, 430.92, 637.23, 677.32],
    "horizontal_strategy": "text",
    "snap_y_tolerance": 6,
}

bbox = (23.64, 117.61, 677.32, 586.69)

input_dir = base_dir / "input"
xlsx_files = input_dir.glob("*.xlsx")
xlsx_file_paths = []
xlsx_file = xlsx_file_paths
for xlsx_file in xlsx_files:
    file_path = str(xlsx_file)
    xlsx_file_paths.append(file_path)

exists = os.path.isfile(input_dir)
initial_version = Path(xlsx_file_paths[1])
updated_version = input_dir / f"{Path(initial_version).stem}_updated.xlsx"
df1 = pd.read_excel(initial_version)
df1.head(10)


# if not exists:
#     writer = pd.ExcelWriter(input_dir, engine="openpyxl")
# else:
#     writer = pd.ExcelWriter(input_dir, mode="a", if_sheet_exists="replace", engine="openpyxl")


pdf_file = input_dir / "EDILOG.pdf"
processed_file = plumber_extract_tables(pdf_file, ts, bbox)
print(f"\n{emoji}  XLSX_FILENAME: {updated_version} {emoji}\n")
with open(pdf_file, 'r') as file:
    loc_cols = ["insurer", "time_entered", "time_made", "policynum", "csrcode", "ccode", "name",
                "renewal"]
    rows = [item[0] for sublist in processed_file.values() for item in sublist]
    df = pd.DataFrame(rows, columns=loc_cols)
    df = df.dropna()
    df_add_col = df.reindex(columns=[*df.columns.tolist(), "pcode", "buscode", "Pulled", "D/L"], fill_value="")
    df_drop = df.drop(columns=["time_entered", "time_made"])
    df2 = df.reindex(columns=["policynum", "ccode", "name", "pcode", "csrcode", "insurer", "buscode", "renewal", "Pulled", "D/L"], fill_value="")

    print(df1)
    print(df2)
    # diff = df2.compare(df1, align_axis=1)

    # sheet_name = datetime.today().strftime("%b %d")
    # writer = pd.ExcelWriter(updated_version, engine="openpyxl")

    # .to_excel(writer, sheet_name="Sheet1", index=False)
    # writer.close()
