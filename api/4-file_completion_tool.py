import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime

base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of Excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# <================================= Formats excel sheet =================================>
df["today"] = datetime.today().strftime("%B %d, %Y")
df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
expiry_date = pd.to_datetime(df["effective_date"]) + pd.offsets.DateOffset(years=1)
df["expiry_date"] = expiry_date.dt.strftime("%B %d, %Y")
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
df[["type", "province", "postal_code"]] = df[["type", "province", "postal_code"]].astype(str).apply(lambda col: col.str.upper())
df[["named_insured", "insurer", "street_address", "city"]] = df[["named_insured", "insurer", "street_address", "city"]].astype(str).apply(lambda x: x.str.title())
df["mailing_address"] = df[["street_address", "city", "province", "postal_code"]].apply(lambda x: ', '.join(x[:-1]) + " " + x[-1:], axis=1)
df["risk_address"] = df["risk_address"].fillna(df["mailing_address"])

# Word Writing Library
def write_to_docx(docx, rows):
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)

    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)


# Directory Paths for each file
filename = {"RENEWAL LETTER": "COVER_Ltr_NEW_RENEWAL_201711.docx"}


for rows in df.to_dict(orient="records"):
    if (rows["type"] == "CONDO RENEWAL"):
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif (rows["type"] == "HOME RENEWAL"):
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif (rows["type"] == "RENTED CONDO RENEWAL"):
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif (rows["type"] == "RENTED DWELLING RENEWAL"):
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif (rows["type"] == "TENANT RENEWAL"):
        write_to_docx(filename["RENEWAL LETTER"], rows)
