import pandas as pd
import re
from helper_funtions import unique_file_name
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of Excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel_path, sheet_name=0, engine="openpyxl")
df = df[df["named_insured"].notna()]

# <================================= Formats excel sheet =================================>
df["today"] = datetime.today().strftime("%B %d, %Y")
df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
expiry_date = pd.to_datetime(df["effective_date"]) + pd.offsets.DateOffset(years=1)
df["expiry_date"] = expiry_date.dt.strftime("%B %d, %Y")
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
df["policy_number"] = df["policy_number"].astype(str).replace(r"\.0$", "", regex=True)
df[["type", "province", "postal_code"]] = df[["type", "province", "postal_code"]].astype(str).apply(
    lambda col: col.str.upper())
df["street_address"] = df["street_address"].astype(str).apply(
    lambda s: s.title() if not re.match(r'\b(\w+)\b', s) else s)
df[["named_insured", "insurer", "city"]] = df[["named_insured", "insurer", "city"]].astype(str).apply(
    lambda x: x.str.title())
df["mailing_address"] = df[["street_address", "city", "province", "postal_code"]].astype(str).apply(
    lambda x: ', '.join(x[:-1]) + " " + x[-1:], axis=1)
df["risk_address"] = df["risk_address"].fillna(df["mailing_address"])

# Word Writing Library
def write_to_docx(docx, rows):
    template_path = base_dir / "input " / "templates"/ docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(unique_file_name(output_path))


# Reads and writes PDF
def write_to_pdf(pdf, dictionary, rows):
    pdf_path = (base_dir / "input " / "templates" / pdf)
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.pdf"
    output_path.parent.mkdir(exist_ok=True)
    with open(unique_file_name(output_path), "wb") as output_stream:
        writer.write(output_stream)


# Directory Paths for each file
filename = {
    "INSURANCE BINDER": "Binder.docx",
    "CANCELLATION RELEASE": "Cancellation Release.docx",
    "GORE RENTED QUESTIONNAIRE": "GORE - Rented Questionnaire.pdf",
    "LETTER OF BROKERAGE": "Letter of Brokerage.docx",
    "FAMILY LOB": "LOB - Family Blank.pdf",
    "RENEWAL LETTER": "Renewal Letter.docx",
    "RENTED INTACT QUESTIONNAIRE": "Rented Intact Questionnaire.docx",
    "REVENUE PROPERTY QUESTIONNAIRE": "Revenue Property Questionnaire.pdf",
    "WAWA MAC AUTHORIZATION FORM": "8003GIS062019MACAuthorizationForm Wawa.pdf",
    "RENTED DWELLING QUESTIONNAIRE": "8186RentedDwellingQuestionnaire0418.pdf",
    "INTACT AUTOMATIC BANK WITHDRAWALS": "Intact withdrawa form.pdf",
}
for rows in df.to_dict(orient="records"):
    if rows["type"] == "CONDO RENEWAL":
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif rows["type"] == "HOME RENEWAL":
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif rows["type"] == "RENTED CONDO RENEWAL":
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif rows["type"] == "RENTED DWELLING RENEWAL":
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif rows["type"] == "TENANT RENEWAL":
        write_to_docx(filename["RENEWAL LETTER"], rows)
    elif rows["type"] == "INSURANCE BINDER":
        write_to_docx(filename["INSURANCE BINDER"], rows)
    elif rows["type"] == "CANCELLATION RELEASE":
        write_to_docx(filename["CANCELLATION RELEASE"], rows)
    elif rows["type"] == "GORE RENTED QUESTIONNAIRE":
        dictionary = {"Applicant / Insured": rows["named_insured"],
                      "Gore Policy #": rows["policy_number"],
                      "Principal Street": rows["street_address"],
                      "Rental Street": rows["risk_address"]
                      }
        write_to_pdf(filename["GORE RENTED QUESTIONNAIRE"], dictionary, rows)
    elif rows["type"] == "LETTER OF BROKERAGE":
        write_to_docx(filename["LETTER OF BROKERAGE"], rows)
    elif rows["type"] == "FAMILY LOB":
        dictionary = {"Name of Insureds": rows["named_insured"],
                      "Address of Insureds": rows["mailing_address"],
                      "Day": rows["effective_date"].split(" ")[1],
                      "Month": rows["effective_date"].split(" ")[0],
                      "Year": rows["effective_date"].split(" ")[2],
                      "Policy Number": rows["policy_number"],
                      "Name 1": rows["named_insured"],
                      "Name 2": rows["additional_insured"],
                      }
        write_to_pdf(filename["FAMILY LOB"], dictionary, rows)
    elif rows["type"] == "RENTED INTACT QUESTIONNAIRE":
        write_to_docx(filename["RENTED INTACT QUESTIONNAIRE"], rows)
    elif rows["type"] == "REVENUE PROPERTY QUESTIONNAIRE":
        dictionary = {"Insured's Name": rows["named_insured"],
                      "Insureds Name": rows["named_insured"],
                      "Policy Number": rows["policy_number"],
                      "Address of Property": rows["risk_address"],
                      "Date Coverage is Required": rows["effective_date"]
                      }
        write_to_pdf(filename["REVENUE PROPERTY QUESTIONNAIRE"], dictionary, rows)
    elif rows["type"] == "WAWA MAC AUTHORIZATION FORM":
        dictionary = {"Policy": rows["policy_number"],
                      "Name": rows["named_insured"],
                      "Address": rows["mailing_address"],
                      "Postal Code": rows["postal_code"],
                      "Please list policy numbers on The MAC Plan": rows["policy_number"],
                      }
        write_to_pdf(filename["WAWA MAC AUTHORIZATION FORM"], dictionary, rows)
    elif rows["type"] == "RENTED DWELLING QUESTIONNAIRE":
        dictionary = {"Insureds Name": rows["named_insured"],
                      "Policy Number": rows["policy_number"],
                      "Address of property": rows["risk_address"],
                      }
        write_to_pdf(filename["RENTED DWELLING QUESTIONNAIRE"], dictionary, rows)
    elif (rows["type"] == "INTACT AUTOMATIC BANK WITHDRAWALS"):
        dictionary = {"Policy Number": rows["policy_number"],
                      "Last Name": rows["named_insured"].split()[-1],
                      "First Name": " ".join(rows["named_insured"].split()[:-1]),
                      "Name of Account Holder": rows["named_insured"],
                      "Province": rows["province"],
                      }
        write_to_pdf(filename["INTACT AUTOMATIC BANK WITHDRAWALS"], dictionary, rows)