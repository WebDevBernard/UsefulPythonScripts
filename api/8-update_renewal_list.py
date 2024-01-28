from helper_funtions import emoji
from pdfplumber_functions import plumber_extract_tables
from pathlib import Path
import pandas as pd

# Updates renewal List from EDI LOG work in progress, EDILOG needs to be placed in input folder to work

# Table Strategy for identifying table, using h-lines or h-text
ts = {
    "vertical_strategy": "explicit",
    "explicit_vertical_lines": [23.64, 50.64, 106.91, 159.36, 217.94, 341.64, 392.67, 430.92, 637.23, 677.32],
    "horizontal_strategy": "text",
    "snap_y_tolerance": 6,
}

bbox = (23.64, 117.61, 677.32, 586.69)

base_dir = Path(__file__).parent.parent
input_directory = Path(base_dir / "input")
pdf_file = input_directory / f"EDILOG.pdf"
processed_file = plumber_extract_tables(pdf_file , ts, bbox)

print(f"\n{emoji}   PDF_FILENAME: {pdf_file} {emoji}\n")
with open(pdf_file, 'w') as file:
    loc_cols = ["insurer", "time_entered", "time_made", "policynum", "csrcode", "ccode", "name",
                "renewal"]
    rows = [item[0] for sublist in processed_file.values() for item in sublist]
    df = pd.DataFrame(rows, columns=loc_cols)
    df = df.dropna()
    print(df)
