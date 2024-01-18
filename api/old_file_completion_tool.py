import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter

# Directory Paths for each file
disclosure_filename = "Disclosure and Waiver.docx"
credit_consent_filename = "Wawa Personal Information and Credit Consent Form 8871.pdf"
questionnaire_filename = ["GORE - Rented Questionnaire.pdf",
                          "Optimum West Rental Q.pdf",
                          "WAWA Rental Condo Questionnaire.pdf",
                          "wawa rented dwelling Q.pdf",
                          "Rented Dwelling Quest INTACT.docx"
                          ]
lob_filename = ["Disclosure and LOB.docx",
                "LOB - Family Blank.pdf"
                ]
misc_filename = ["NCOL Letter.docx",
                 "Express Consent.docx",
                 "Coverage Deletion Request.docx",
                 "Commercial Consent Form.docx"
                 ]
ncol_filename = "NCOL Letter.docx"
binder_binder_fee = ["Binder Fee Invoice - Cedar.pdf",
                     "Binder.docx"
                     ]
wawa_banking_info = "Wawa monthly payplan Authorization Form, Form 8003GIS.pdf"
cancel_letter = "Cancellation-Mid-Term-or-Flat Letter.docx"
base_dir = Path(__file__).parent.parent
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

# Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Format date to MMM DD, YYYY
df["today"] = datetime.today().strftime("%B %d, %Y")
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")

# Remove white spaces, capitalize first letter of each word, and format effective date time
def namesFormatter(df):
    for i in df.columns:
        if df[i].dtype == 'object':
            df[i] = df[i].str.strip()
        if df["policy_number"].dtype == 'object':
            df["policy_number"] = df["policy_number"].astype("string")
        if df["type"].dtype == 'object':
            df["type"] = df["type"].str.upper()
        if df["additional"].dtype == 'object':
            df["additional"] = df["additional"].str.upper()
        if df["insured_name"].dtype == 'object':
            df["insured_name"] = df["insured_name"].str.title()
        if df["additional_insured"].dtype == 'object':
            df["additional_insured"] = df["additional_insured"].str.title()
        if df["insurer"].dtype == 'object':
            df["insurer"] = df["insurer"].str.title()
        if df["broker_name"].dtype == 'object':
            df["broker_name"] = df["broker_name"].str.title()
        else:
            pass

namesFormatter(df)

# Checks if there is additonal_insured
def insuredNames(rows):
    if (pd.isnull(rows["additional_insured"])):
        return rows["insured_name"]
    return rows["insured_name"] + " & " + rows['additional_insured']

# Check family if addional insured but return only empty string
def additionalInsured(rows):
    if (pd.isnull(rows["additional_insured"])):
        return ""
    return rows["additional_insured"]

# Checks if no risk address, use mailing address as risk address
def riskAddress(rows):
    if (pd.isnull(rows["risk_address"])):
        return rows["mailing_address"]
    return rows["risk_address"]

# Checks if 2nd location, use mailing address as second location
def secondLocation(rows):
    if (pd.isnull(rows["secondary_location"])):
        return rows["mailing_address"]
    return rows["secondary_location"]


# Checks if there is a mailing address
def checkMailingAddress(rows):
    if (pd.isnull(rows["mailing_address"])):
        return ""
    return rows["mailing_address"]


# Checks if there is a policy_number
def checkPolicyNumber(rows):
    if (pd.isnull(rows["policy_number"])):
        return ""
    return rows["policy_number"]


# Checks if there is an effective date
def checkEffectiveDate(rows):
    if (pd.isnull(rows["effective_date"])):
        return ""
    return rows["effective_date"]


# Reads and writes PDF
def writeToPdf(pdf, dictionary, rows):
    pdf_path = base_dir / "input" / pdf
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for pageNum in range(reader.numPages):
        page = reader.getPage(pageNum)
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} {pdf}"
    output_path.parent.mkdir(exist_ok=True)
    with open(output_path, "wb") as output_stream:
        writer.write(output_stream)

# Write to Docx
def writeToDocx(docx, rows):
    template_path = base_dir / "input" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} {docx}"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)


# Write to LOB Docx
def writeToFamilyDocx(docx, rows):
    template_path = base_dir / "input" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    if (rows["insurer"] == "Family"):
        output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} Disclosure Notice.docx"
    else:
        output_path = output_dir / f"{insuredNames(rows)}" / f"{insuredNames(rows)} - {checkPolicyNumber(rows)} {docx}"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)


for rows in df.to_dict(orient="records"):
    if (rows["type"] == "LOB"):
        # Make Disclosure and LOB.docx
        writeToFamilyDocx(lob_filename[0], rows)
        # Make LOB - Family Blank.pdf
        if (rows["insurer"] == "Family"):
            dictionary = {"Name of Insureds": insuredNames(rows),
                          "Address of Insureds": checkMailingAddress(rows),
                          "Day": checkEffectiveDate(rows).split(" ")[1],
                          "Month": checkEffectiveDate(rows).split(" ")[0],
                          "Year": checkEffectiveDate(rows).split(" ")[2],
                          "Policy Number": checkPolicyNumber(rows),
                          "Name 1": rows["insured_name"],
                          "Name 2": additionalInsured(rows)
                          }
            writeToPdf(lob_filename[1], dictionary, rows)
    if (rows["type"] == "NEW BUSINESS"):
        writeToDocx(disclosure_filename, rows)
        # Makes Wawa Personal Information and Credit Consent Form 8871
        if (rows["insurer"] == "Wawanesa"):
            dictionary = {"Policy  Submission Numbers": checkPolicyNumber(rows),
                          "Insureds Name": insuredNames(rows),
                          }
            writeToPdf(credit_consent_filename, dictionary, rows)
    if (rows["additional"] == "NCOL"):
        writeToDocx(ncol_filename, rows)
    if (rows["additional"] == "EXPRESS CONSENT"):
        writeToDocx(misc_filename[1], rows)
    if (rows["type"] == "CHANGE"):
        writeToDocx(misc_filename[2], rows)
    if (rows["type"] ==  "COMMERCIAL"):
        writeToDocx(misc_filename[3], rows)
    if (rows["type"] == "CANCEL"):
        writeToDocx(cancel_letter, rows)
    if (rows["type"] == "BINDER"):
        dictionary = {'Effective Date': checkEffectiveDate(rows),
                      'Policy Number': checkPolicyNumber(rows),
                      }
        writeToPdf(binder_binder_fee[0], dictionary, rows)
        writeToDocx(binder_binder_fee[1], rows)
    if (rows["type"] == "LOB" or rows["type"] == "NEW BUSINESS" or rows["type"] == "CHANGE"):
        if pd.notnull(rows["risk_address"]):
            # Make GORE - Rented Questionnaire
            if (rows["insurer"] == "Gore Mutual"):
                dictionary = {"Applicant / Insured": insuredNames(rows),
                              "Gore Policy #": checkPolicyNumber(rows),
                              "Principal Street": checkMailingAddress(rows),
                              "Rental Street": riskAddress(rows)
                              }
                writeToPdf(questionnaire_filename[0], dictionary, rows)
            # Make Questionnaire - Optimum West Rental Q
            if (rows["insurer"] == "Optimum West"):
                dictionary = {"Policy_Number[0]": checkPolicyNumber(rows),
                              "Applicant_Insured[0]": insuredNames(rows),
                              "Rental_Location_Address[0]": riskAddress(rows),
                              }
                writeToPdf(questionnaire_filename[1], dictionary, rows)
            # Make Questionnaire - WAWA Rental Condo Questionnaire
            if (rows["insurer"] == "Wawanesa" and rows["additional"] != "WAWANESA REVENUE"):
                dictionary = {"Insured's Name": insuredNames(rows),
                              "Insureds Name": insuredNames(rows),
                              "Policy Number": checkPolicyNumber(rows),
                              "Address of Property": riskAddress(rows),
                              "Date Coverage is Required": checkEffectiveDate(rows),
                              }
                writeToPdf(questionnaire_filename[2], dictionary, rows)
            # Make Questionnaire - wawa rented dwelling Q
            if (rows["insurer"] == "Wawanesa" and rows["additional"] == "WAWANESA REVENUE"):
                dictionary = {"Insured's Name": insuredNames(rows),
                              "Policy Number": checkPolicyNumber(rows),
                              "Address of Property": riskAddress(rows),
                              }
                writeToPdf(questionnaire_filename[3], dictionary, rows)
            # Make Wawa Banking info Form
            if (rows["insurer"] == "Wawanesa" and rows["additional"] == "WAWANESA MONTHLY"):
                dictionary = {"Name": insuredNames(rows),
                              "Address": riskAddress(rows),
                              "Policy": checkPolicyNumber(rows)
                              }
                writeToPdf(wawa_banking_info, dictionary, rows)
            # Make Questionnaire - Rented Dwelling Quest INTACT
            if (rows["insurer"] == "Intact"):
                writeToDocx(questionnaire_filename[4], rows)

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