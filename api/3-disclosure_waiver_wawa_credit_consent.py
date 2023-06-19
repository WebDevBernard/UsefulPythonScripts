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

writer = PdfWriter()

#Reads and writes PDF
def writeToPdf(pdf, dictionary, rows):
  pdf_path = base_dir / "input" / pdf
  reader = PdfReader(pdf_path)
  for pageNum in range(reader.numPages):
    page = reader.getPage(pageNum)
    writer.add_page(page)
  writer.updatePageFormFieldValues(
  writer.getPage(0), dictionary
  )
  output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {pdf}"
  output_path.parent.mkdir(exist_ok=True)
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)
    
# Loop generates LOB and DN
for rows in df.to_dict(orient="records"):
    doc = DocxTemplate(LOB_template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {docx_filename}"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)
    pdf_filename = ["GORE - Rented Questionnaire.pdf", 
                  "Optimum West Rental Q.pdf", 
                  "WAWA Rental Condo Questionnaire.pdf",
                  "wawa rented dwelling Q.pdf",
                  "Wawa Personal Information and Credit Consent Form 8871.pdf"
                  ]
    if (rows["insurer"] == "Wawanesa"):
        dictionary = {"Policy  Submission Numbers": rows["policy_number"],
                    "Insureds Name": rows["insured_name"],
                    }
        writeToPdf(pdf_filename[4], dictionary, rows)
    if (rows["insurer"] == "Intact" and rows["type"] == "Rental"):
        doc = DocxTemplate(LOB_template_path)
        doc.render(rows)
        output_path = output_dir / f"{rows['insured_name']}" / f"{rows['insured_name']} - {docx_filename}"
        output_path.parent.mkdir(exist_ok=True)
        doc.save(output_path)
    if (rows["insurer"] == "Gore Mutual" and rows["type"] == "Rental"):
      dictionary = {"Applicant / Insured": rows["insured_name"],
                "Gore Policy #": rows["policy_number"],
                "Principal Street": rows["mailing_address"],
                "Rental Street": rows["risk_address"]
                  }
      writeToPdf(pdf_filename[0], dictionary, rows)
    if (rows["insurer"] == "Optimum" and rows["type"] == "Rental"):
      dictionary = {"Policy_Number[0]": rows["policy_number"],
                  "Applicant_Insured[0]": rows["insured_name"],
                  "Rental_Location_Address[0]": rows["risk_address"],
                  }
      writeToPdf(pdf_filename[1], dictionary, rows)
    if (rows["insurer"] == "Wawanesa" and rows["type"] == "Rental"):
      dictionary = {"Insureds Name": rows["insured_name"],
                  "Policy Number": rows["policy_number"],
                  "Address of Property": rows["risk_address"],
                  "Date Coverage is Required": rows["effective_date"],
                  }
      writeToPdf(pdf_filename[2], dictionary, rows)
    if (rows["insurer"] == "Wawanesa" and rows["type"] == "Revenue"):
      dictionary = {"Insured's Name": rows["insured_name"],
                  "Policy Number": rows["policy_number"],
                  "Address of Property": rows["risk_address"],
                  }
      writeToPdf(pdf_filename[3], dictionary, rows)