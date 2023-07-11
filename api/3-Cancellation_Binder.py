import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

#Directory Paths for each file
binder_binder_fee = [ "Binder Fee Invoice - Cedar.pdf",
                      "Binder.docx"
                    ]
cancel_letter = "Cancellation-Mid-Term-or-Flat Letter.docx"
base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output" # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet3")

#Format date to MMM DD, YYYY
df["today"] = datetime.today().strftime("%B %d, %Y")
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")

#Remove white spaces, capitalize first letter of each word, and format effective date time
def namesFormatter(df):
    for i in df.columns:
      if df[i].dtype == 'object':
        df[i] = df[i].str.strip()
      if df["policy_number"].dtype == 'object':
        df["policy_number"] = df["policy_number"].astype("string")
      if df["type"].dtype == 'object':
        df["type"] = df["type"].str.title()
      if df["insured_name"].dtype == 'object':  
        df["insured_name"] = df["insured_name"].str.title()
      if df["insurer"].dtype == 'object':
        df["insurer"] = df["insurer"].str.title()
      if df["broker_name"].dtype == 'object':
        df["broker_name"] = df["broker_name"].str.title()
      else:
        pass
namesFormatter(df)

#Checks if there is additonal_insured
def insuredNames(rows):
  if (pd.isnull(rows["additional_insured"])):
    return rows["insured_name"]
  return rows["insured_name"] + " & " + rows['additional_insured']

#Checks if no risk address, use mailing address as risk address
def riskAddress(rows):
  if (pd.isnull(rows["risk_address"])):
    return rows["mailing_address"]
  return rows["risk_address"]

#Checks if there is a policy_number
def checkPolicyNumber(rows):
  if (pd.isnull(rows["policy_number"])):
    return ""
  return rows["policy_number"]

#Checks if there is a policy_number
def checkClientCode(rows):
  if (pd.isnull(rows["client_code"])):
    return ""
  return rows["client_code"]

#Checks if there is an effective date
def checkEffectiveDate(rows):
  if (pd.isnull(rows["effective_date"])):
    return ""
  return rows["effective_date"]

#Reads and writes PDF
def writeToPdf(pdf, dictionary, rows):
  pdf_path = base_dir / "input" / pdf
  reader = PdfReader(pdf_path)
  writer = PdfWriter()
  for pageNum in range(reader.numPages):
    page = reader.getPage(pageNum)
    writer.add_page(page)
  writer.updatePageFormFieldValues(
  writer.getPage(0), dictionary
  )
  output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {pdf}"
  output_path.parent.mkdir(exist_ok=True)
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)
    
#Write to Docx    
def writeToDocx(docx, rows):
  template_path = base_dir / "input" / docx
  doc = DocxTemplate(template_path)
  doc.render(rows)
  output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {docx}"
  output_path.parent.mkdir(exist_ok=True)
  doc.save(output_path)    

for rows in df.to_dict(orient="records"):
  # Make Cancel/Lapse Letter
  if (rows["type"] == "Cancel" or rows["type"] == "Lapse"):
    writeToDocx(cancel_letter, rows)
  #Binder Invoice and Binder
  if (rows["type"] == "Binder"):
    dictionary = {'Effective Date': checkEffectiveDate(rows),
                  'Policy Number' : checkPolicyNumber(rows),
                  'Account Number' : checkClientCode(rows),
                }
    writeToPdf(binder_binder_fee[0], dictionary, rows)
    writeToDocx(binder_binder_fee[1], rows)