import pandas as pd 
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

#Directory Paths for each file
lob_filename = ["Disclosure and LOB.docx",
                "LOB - Family Blank.pdf"
                ]
questionnaire_filename = ["GORE - Rented Questionnaire.pdf", 
                          "Optimum West Rental Q.pdf", 
                          "WAWA Rental Condo Questionnaire.pdf",
                          "wawa rented dwelling Q.pdf",
                          "Rented Dwelling Quest INTACT.docx"
                          ]
base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output" # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet1")

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
  output_path = output_dir / f"{insuredNames(rows)} - {insuredNames(rows)} {pdf}"
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)
    
#Write to Docx    
def writeToDocx(docx, rows):
  template_path = base_dir / "input" / docx
  doc = DocxTemplate(template_path)
  doc.render(rows)
  if (rows["insurer"] == "Family"):
    output_path = output_dir / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} Disclosure Notice.docx"
  else:
    output_path = output_dir / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} {docx}"
  output_path.parent.mkdir(exist_ok=True)
  doc.save(output_path)

for rows in df.to_dict(orient="records"):
  #Make Disclosure and LOB.docx
  writeToDocx(lob_filename[0],rows)
  #Make LOB - Family Blank.pdf
  if (rows["insurer"] == "Family"):
    dictionary = {"Name of Insureds": insuredNames(rows),
                  "Address of Insureds": rows["mailing_address"],
                  # "Day": rows["effective_date"].split(" ")[1],
                  # "Month": rows["effective_date"].split(" ")[0],
                  # "Year": rows["effective_date"].split(" ")[2],
                  "Policy Number": checkPolicyNumber(rows),
                  }
    writeToPdf(lob_filename[1], dictionary, rows)
  if pd.notnull(rows["risk_address"]):
    #Make GORE - Rented Questionnaire
    if (rows["insurer"] == "Gore Mutual"):
      dictionary = {"Applicant / Insured": insuredNames(rows),
                    "Gore Policy #": checkPolicyNumber(rows),
                    "Principal Street": rows["mailing_address"],
                    "Rental Street": riskAddress(rows)
                  }
      writeToPdf(questionnaire_filename[0], dictionary, rows)
    #Make Questionnaire - Optimum West Rental Q
    if (rows["insurer"] == "Optimum West"):
      dictionary = {"Policy_Number[0]": checkPolicyNumber(rows),
                    "Applicant_Insured[0]": insuredNames(rows),
                    "Rental_Location_Address[0]": riskAddress(rows),
                    }
      writeToPdf(questionnaire_filename[1], dictionary, rows)
    #Make Questionnaire - WAWA Rental Condo Questionnaire
    if (rows["insurer"] == "Wawanesa" and rows["type"] != "Revenue"):
      dictionary = {"Insureds Name": insuredNames(rows),
                    "Policy Number": checkPolicyNumber(rows),
                    "Address of Property": riskAddress(rows),
                    "Date Coverage is Required": rows["effective_date"],
                    }
      writeToPdf(questionnaire_filename[2], dictionary, rows)
    #Make Questionnaire - wawa rented dwelling Q 
    if (rows["insurer"] == "Wawanesa" and rows["type"] == "Revenue"):
      dictionary = {"Insured's Name": insuredNames(rows),
                    "Policy Number": checkPolicyNumber(rows),
                    "Address of Property": riskAddress(rows),
                    }
      writeToPdf(questionnaire_filename[3], dictionary, rows)   
    #Make Questionnaire - Rented Dwelling Quest INTACT 
    if (rows["insurer"] == "Intact"):
      writeToDocx(questionnaire_filename[4], rows)