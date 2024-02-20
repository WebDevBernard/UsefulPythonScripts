import pandas as pd
import os
from pathlib import Path
from coordinates import base_dir

input_dir = base_dir / "input"
xlsx_files = Path(input_dir).glob("*.xlsx")
xls_files = Path(input_dir).glob("*.xls")
files = list(xlsx_files) + list(xls_files)
excel_path = list(files)[0]
output_path = base_dir / "output" / f"{Path(excel_path).stem}.xlsx"


def make_df_from_excel():
    if Path(excel_path).suffix == ".XLS":
        return pd.read_excel(excel_path, engine="xlrd")
    elif Path(excel_path).suffix == ".XLSX":
        return pd.read_excel(excel_path, engine="openpyxl")


def sort_renewal_list():
    df = make_df_from_excel()
    column_list = ["policynum", "ccode", "name", "pcode", "csrcode", "insurer", "buscode", "renewal", "Pulled", "D/L"]
    df = df.reindex(columns=column_list)
    df = df.drop_duplicates(subset=["policynum"], keep=False)
    df.sort_values(["insurer", "renewal", "name"], ascending=[True, True, True], inplace=True)
    list_with_spaces = []
    for x, y in df.groupby('insurer', sort=False):
        list_with_spaces.append(y)
        list_with_spaces.append(pd.DataFrame([[float('NaN')] * len(y.columns)], columns=y.columns))
    df = pd.concat(list_with_spaces, ignore_index=True).iloc[:-1]
    print(df)
    if not os.path.isfile(output_path):
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    else:
        writer = pd.ExcelWriter(output_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
    df.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.close()
