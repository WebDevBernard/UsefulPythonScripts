import pandas as pd  # pip install pandas openpyxl
from datetime import datetime, timedelta
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / "WAWA Rental Condo Questionnaire.pdf"
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Formats dates to MMM DD, YYYY
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")
df["today"] = datetime.today().strftime("%B %d, %Y")

reader = PdfReader(pdf_path)
writer = PdfWriter()
page = reader.pages[0]
writer.add_page(page)

dictionary = {"Insureds Name": df["insured_name"].values[0],
              "Policy Number": df["policy_number"].values[0],
              "Address of Property": df["risk_address"].values[0],
              "Date Coverage is Required": df["effective_date"].values[0],   
              "Date": df["today"].values[0],   
              "Date_2": df["today"].values[0],   
              }

dictionary_2 = {"DATE (mm/dd/yyyy)": df["today"].values[0],
              }

writer.updatePageFormFieldValues(
    writer.getPage(0), dictionary
)

for record in df.to_dict(orient="records"):
  output_path = output_dir / f"{df['insured_name'].values[0]} - WAWA Rental Condo Questionnaire.pdf" 
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)
