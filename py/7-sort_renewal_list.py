import pandas as pd
import os
from pathlib import Path
from helper_fn import base_dir

input_dir = base_dir / "input"
files = Path(input_dir).glob("*.xls")
excel_path = list(files)[0]
output_path = base_dir / "output" / f"{Path(excel_path).stem}.xlsx"
df = pd.read_excel(excel_path, engine="xlrd")
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


if not os.path.isfile(output_path):
    writer = pd.ExcelWriter(output_path, engine="openpyxl")
else:
    writer = pd.ExcelWriter(output_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
df.to_excel(writer, sheet_name="Sheet1", index=False)
writer.close()
