import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

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
df[["insurer", "street_address", "city"]] = df[["insurer", "street_address", "city"]].astype(str).apply(lambda x: x.str.title())
df["mailing_address"] = df[["street_address", "city", "province", "postal_code"]].apply(lambda x: ', '.join(x[:-1]) + " " + x[-1:], axis=1)
df["risk_address"] = df["risk_address"].fillna(df["mailing_address"])


# PDF Writing Library
def write_to_pdf(pdf, dictionary, rows):
    pdf_path = base_dir / "input" / pdf
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for pageNum in range(reader.numPages):
        page = reader.getPage(pageNum)
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.pdf"
    output_path.parent.mkdir(exist_ok=True)
    with open(output_path, "wb") as output_stream:
        writer.write(output_stream)


# Word Writing Library
def write_to_docx(docx, rows):
    template_path = base_dir / "input" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)


# Directory Paths for each file
filename = {"CONDO RENEWAL": "COVER_Ltr_NEW_Condo_201711.docx",
            "HOME RENEWAL": "COVER_Ltr_NEW_Homeowners_201711.docx",
            "RENTED CONDO RENEWAL": "COVER_Ltr_NEW_RentedCondo_201711.docx",
            "RENTED DWELLING RENEWAL": "COVER_Ltr_NEW_RentedDwelling_201711.docx",
            "TENANTS RENEWAL": "COVER_Ltr_NEW_tenants_201711.docx"}


for rows in df.to_dict(orient="records"):
    if (rows["type"] == "CONDO RENEWAL"):
        write_to_docx(filename["CONDO RENEWAL"], rows)
    elif (rows["type"] == "HOME RENEWAL"):
        write_to_docx(filename["HOME RENEWAL"], rows)
    elif (rows["type"] == "RENTED CONDO RENEWAL"):
        write_to_docx(filename["RENTED CONDO RENEWAL"], rows)
    elif (rows["type"] == "RENTED DWELLING RENEWAL"):
        write_to_docx(filename["RENTED DWELLING RENEWAL"], rows)
    elif (rows["type"] == "TENANTS RENEWAL"):
        write_to_docx(filename["TENANTS RENEWAL"], rows)
