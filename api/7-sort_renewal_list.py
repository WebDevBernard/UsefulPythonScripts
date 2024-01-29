import pandas as pd
import os
from pathlib import Path
from datetime import datetime
from helper_funtions import base_dir

input_dir = base_dir / "input"  # name of output folder
files = Path(input_dir).glob("*.xlsx")
file_paths = []

for file in files:
    file_path = str(file)
    file_paths.append(file_path)
# Open only first xlsx file in folder

excel_path = file_paths[0]
output_path = input_dir / f"{Path(excel_path).stem}_sorted_renewal_list.xlsx"
df = pd.read_excel(excel_path, sheet_name=0, engine="openpyxl")
exists = os.path.isfile(output_path)

column_list = ["policynum", "ccode", "name", "pcode", "csrcode", "insurer", "buscode", "renewal", "Pulled", "D/L"]
df = df.reindex(columns=column_list)
df.sort_values(["insurer", "renewal", "name"], ascending=[True, True, True], inplace=True)
list_with_spaces = []
for x, y in df.groupby('insurer', sort=False):
    list_with_spaces.append(y)
    list_with_spaces.append(pd.DataFrame([[float('NaN')] * len(y.columns)], columns=y.columns))
df = pd.concat(list_with_spaces, ignore_index=True).iloc[:-1]


def highlight_duplicates(s, other_column):
    is_duplicate = (other_column != 'GLA') & (other_column != 'EQB') & s.duplicated(keep=False) & (~s.isna())
    return ['background-color: LightGreen' if v else '' for v in is_duplicate]

duplicated = df['ccode'].duplicated()
df = df.style.apply(highlight_duplicates, subset=["ccode"], other_column=df["buscode"])

emoji = "\U0001F923\U0001F923\U0001F923\U0001F923\U0001F923"
print(f"\n{emoji}    XlSX_FILENAME: {output_path.stem} {emoji}\n")

if not exists:
    writer = pd.ExcelWriter(output_path, engine="openpyxl")
else:
    writer = pd.ExcelWriter(output_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
sheet_name = datetime.today().strftime("%b %d")
df.to_excel(writer, sheet_name="Sheet1", index=False)
writer.close()
