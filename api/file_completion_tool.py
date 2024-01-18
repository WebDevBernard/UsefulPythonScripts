import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import jinja2
import re

base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of Excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Directory Paths for each file
filename = {"condo_renewal_letter": "RENEWAL_Condo_201711.docx"
            }
# <================================= Formats excel sheet =================================>
df["today"] = datetime.today().strftime("%B %d, %Y")
df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
expiry_date = pd.to_datetime(df["effective_date"]) + pd.offsets.DateOffset(years=1)
df["expiry_date"] = expiry_date.dt.strftime("%B %d, %Y")
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
df[["type", "additional", "province", "postal_code"]] = df[["type", "additional", "province", "postal_code"]].astype(str).apply(lambda col: col.str.upper())
df[["insured_name", "employee_name", "insurer", "city"]] = df[["insured_name", "employee_name", "insurer", "city"]].astype(str).apply(
    lambda col: col.str.title())
# <================================= Replace NAN risk address =================================>
df["risk_address"] = df["risk_address"].fillna(df["street_address"] + ", " + df["city"] + ", " + df["province"] + " " + df["postal_code"])

# PDF Writing Library
def write_to_pdf(pdf, dictionary, rows):
    pdf_path = base_dir / "input" / pdf
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for pageNum in range(reader.numPages):
        page = reader.getPage(pageNum)
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    output_path = output_dir / f"{rows["insured_name"]}" / f"{rows["insured_name"]} {rows["type"].title()} .pdf"
    output_path.parent.mkdir(exist_ok=True)
    with open(output_path, "wb") as output_stream:
        writer.write(output_stream)

# Word Writing Library
def write_to_docx(docx, rows):
    template_path = base_dir / "input" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["insured_name"]}" / f"{rows["insured_name"]} {rows["type"].title()} .docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)

for rows in df.to_dict(orient="records"):
    rows["mailing_address"] = f"{rows["street_address"]}, {rows["city"]}, {rows["province"]} {rows["postal_code"]}"
    if (rows["type"] == "CONDO RENEWAL LETTER"):
        write_to_docx(filename["condo_renewal_letter"], rows)
