from file_writing_fn import plumber_extract_tables
import pandas as pd
from pathlib import Path
import os
from helper_fn import base_dir


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
xlsx_file = xlsx_file_paths[1]
output_path = input_dir / f"{Path(xlsx_file).stem}_updated.xlsx"

for xlsx_file in xlsx_files:
    file_path = str(xlsx_file)
    xlsx_file_paths.append(file_path)

df1 = pd.read_excel(xlsx_file, engine="openpyxl")
df1.head(10)

pdf_file = input_dir / "EDILOG.pdf"
processed_file = plumber_extract_tables(pdf_file, ts, bbox)
with open(pdf_file, 'r') as file:
    loc_cols = ["insurer", "time_entered", "time_made", "policynum", "csrcode", "ccode", "name",
                "renewal"]
    rows = [item[0] for sublist in processed_file.values() for item in sublist]
    df2 = pd.DataFrame(rows, columns=loc_cols)
    df2 = df2.dropna()
    df_add_col = df2.reindex(columns=[*df2.columns.tolist(), "pcode", "buscode", "Pulled", "D/L"], fill_value="")
    df_drop = df2.drop(columns=["time_entered", "time_made"])
    df2 = df2.reindex(columns=["policynum", "ccode", "name", "pcode", "csrcode", "insurer", "buscode", "renewal",
                              "Pulled", "D/L"], fill_value="")
    merged_df = pd.merge(df1, df2[['policynum', 'D/L']], on='policynum', how='left', suffixes=('_df1', '_df2'))
    merged_df['D/L_df1'] = merged_df.apply(lambda row: row['D/L_df2'] if pd.notnull(row['D/L_df2']) else row['D/L_df1'],
                                           axis=1)
    df1 = merged_df.drop(['D/L_df2'], axis=1)

    def highlight_changes(val):
        color = '' if val != val else 'background-color: LightGreen'
        return color

    df1 = df1.style.map(highlight_changes, subset=['D/L'])

    if not os.path.isfile(output_path):
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
    else:
        writer = pd.ExcelWriter(output_path, mode="a", if_sheet_exists="replace", engine="openpyxl")

    df1.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.close()
