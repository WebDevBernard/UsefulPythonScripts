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
df[["type", "additional"]] = df[["type", "additional"]].astype(str).apply(lambda col: col.str.upper())
df[["insured_name", "insurer", "employee_name"]] = df[["insured_name", "insurer", "employee_name"]].astype(str).apply(
    lambda col: col.str.title())
# <================================= Replace NAN risk address =================================>
df["risk_address"] = df["risk_address"].fillna(df["mailing_address"])


# # <================================= Address Formatter WIP =================================>
# city_regex = r"(?i)avenue|ave|boulevard|blvd|court|crt|ct|highway|hwy|street|st"
# postal_code_regex = r"(?i)[ABCEGHJ-NPRSTVXY][0-9][ABCEGHJ-NPRSTV-Z][ -]?[0-9][ABCEGHJ-NPRSTV-Z][0-9]"
# province_regex = r"(?i)\b(NL|PE|NS|NB|QC|ON|MB|SK|AB|BC|YT|NT|NU)"
# def regex(string):
#     address_list = []
#     address_list.append(re.split(city_regex, string)[0])
#     address_list.append(re.findall(city_regex, string)[0])
#     address_list.append(re.findall(province_regex, string)[0])
#     address_list.append(re.findall(postal_code_regex, string)[0])
#     print(address_list)
# jinja2.filters.FILTERS["regex"] = regex
# for index, row in df.iterrows():
#     regex(row["mailing_address"])
#     re.findall(postal_code_regex, row["mailing_address"])
# <================================= Address Formatter WIP =================================>

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
    if (rows["type"] == "CONDO RENEWAL LETTER"):
        write_to_docx(filename["condo_renewal_letter"], rows)
