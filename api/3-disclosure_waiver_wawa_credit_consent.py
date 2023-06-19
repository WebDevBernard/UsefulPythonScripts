import pandas as pd  # pip install pandas openpyxl
from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl
from datetime import datetime, timedelta
from PyPDF2 import PdfReader, PdfWriter

#Directory Paths for each file
pdf_filename = "Wawa Personal Information and Credit Consent Form 8871.pdf" #filename
docx_filename = "Disclosure and Waiver.docx" #filename

base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / pdf_filename
LOB_template_path = base_dir / "input" / docx_filename   # name of LOB doc
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output" # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Formats dates to MMM DD, YYYY
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")
thirty_before_effective = pd.to_datetime(df["effective_date"], format="%B %d, %Y") - timedelta(days=30)
df["thirty_before_effective"] = thirty_before_effective.dt.strftime("%B %d, %Y")
df["today"] = datetime.today().strftime("%B %d, %Y")

#PyPDF2 Magic
reader = PdfReader(pdf_path)
writer = PdfWriter()
for pageNum in range(reader.numPages):
  page = reader.getPage(pageNum)
  writer.add_page(page)

#PyPDF2 Code Duplication I will fix later

# 1 GORE QUESTIONNAIRE
pdf_filename1 = "GORE - Rented Questionnaire.pdf" #filename
pdf_path1 = base_dir / "input" / pdf_filename1
reader1 = PdfReader(pdf_path1)
writer1 = PdfWriter()
for pageNum in range(reader1.numPages):
  page = reader1.getPage(pageNum)
  writer1.add_page(page)

# 2 OPTIMUM QUESTIONNAIRE
pdf_filename2 = "Optimum West Rental Q.pdf" #filename
pdf_path2 = base_dir / "input" / pdf_filename2
reader2 = PdfReader(pdf_path2)
writer2 = PdfWriter()
for pageNum in range(reader2.numPages):
  page = reader2.getPage(pageNum)
  writer2.add_page(page)

# 3 WAWA RENTAL CONDO QUESTIONNAIRE
pdf_filename3 = "WAWA Rental Condo Questionnaire.pdf" #filename
pdf_path3 = base_dir / "input" / pdf_filename3
reader3 = PdfReader(pdf_path3)
writer3 = PdfWriter()
for pageNum in range(reader3.numPages):
  page = reader3.getPage(pageNum)
  writer3.add_page(page)

# 4 WAWA RENTAL DWELLING QUESTIONNAIRE
pdf_filename4 = "wawa rented dwelling Q.pdf" #filename
pdf_path4 = base_dir / "input" / pdf_filename4
reader4 = PdfReader(pdf_path4)
writer4 = PdfWriter()
for pageNum in range(reader4.numPages):
  page = reader4.getPage(pageNum)
  writer4.add_page(page)
    
# Loop generates LOB and DN
for rows in df.to_dict(orient="records"):
    doc = DocxTemplate(LOB_template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {docx_filename}"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)
    if (rows["insurer"] == "Wawanesa"):
        dictionary = {"Policy  Submission Numbers": rows["policy_number"],
                    "Insureds Name": rows["insured_name"],
                    }
        writer.updatePageFormFieldValues(
        writer.getPage(0), dictionary
        )
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf_filename}"
        output_path.parent.mkdir(exist_ok=True)
        with open(output_path, "wb") as output_stream:
            writer.write(output_stream)
    if (rows["insurer"] == "Intact" and rows["type"] == "Rental"):
        doc = DocxTemplate(LOB_template_path)
        doc.render(rows)
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {docx_filename}"
        output_path.parent.mkdir(exist_ok=True)
        doc.save(output_path)
    if (rows["insurer"] == "Gore Mutual" and rows["type"] == "Rental"):
        # Dictionary
        dictionary = {"Applicant / Insured": rows["insured_name"],
                    "Gore Policy #": rows["policy_number"],
                    "Principal Street": rows["mailing_address"],
                    "Rental Street": rows["risk_address"]
                    }
        writer1.updatePageFormFieldValues(
        writer1.getPage(0), dictionary
        )
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf_filename1}"
        output_path.parent.mkdir(exist_ok=True)
        with open(output_path, "wb") as output_stream:
            writer1.write(output_stream)
    if (rows["insurer"] == "Optimum" and rows["type"] == "Rental"):
        # Dictionary
        dictionary = {"Policy_Number[0]": rows["policy_number"],
                    "Applicant_Insured[0]": rows["insured_name"],
                    "Rental_Location_Address[0]": rows["risk_address"],
                    }
        writer2.updatePageFormFieldValues(
        writer2.getPage(0), dictionary
        )
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf_filename2}"
        output_path.parent.mkdir(exist_ok=True)

        with open(output_path, "wb") as output_stream:
            writer2.write(output_stream)
    if (rows["insurer"] == "Wawanesa" and rows["type"] == "Rental"):
        # Dictionary
        dictionary = {"Insureds Name": rows["insured_name"],
                    "Policy Number": rows["policy_number"],
                    "Address of Property": rows["risk_address"],
                    "Date Coverage is Required": rows["effective_date"],
                    }
        writer3.updatePageFormFieldValues(
        writer3.getPage(0), dictionary
        )
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf_filename3}"
        output_path.parent.mkdir(exist_ok=True)
        with open(output_path, "wb") as output_stream:
            writer3.write(output_stream)
    if (rows["insurer"] == "Wawanesa" and rows["type"] == "Revenue"):
        # Dictionary
        dictionary = {"Insured's Name": rows["insured_name"],
                    "Policy Number": rows["policy_number"],
                    "Address of Property": rows["risk_address"],
                    }
        writer4.updatePageFormFieldValues(
        writer4.getPage(0), dictionary
        )
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf_filename4}"
        output_path.parent.mkdir(exist_ok=True)
        with open(output_path, "wb") as output_stream:
            writer4.write(output_stream)