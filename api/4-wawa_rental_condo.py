import pandas as pd  # pip install pandas openpyxl
from datetime import datetime, timedelta
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

pdf_filename = "WAWA Rental Condo Questionnaire.pdf" #filename
base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / pdf_filename
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# Formats dates to MMM DD, YYYY
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")
df["today"] = datetime.today().strftime("%B %d, %Y")

# Dictionary
dictionary = {"Insureds Name": df["insured_name"].values[0],
              "Policy Number": df["policy_number"].values[0],
              "Address of Property": df["risk_address"].values[0],
              "Date Coverage is Required": df["effective_date"].values[0],
              }

reader = PdfReader(pdf_path)
writer = PdfWriter()
for pageNum in range(reader.numPages):
  page = reader.getPage(pageNum)
  writer.add_page(page)

writer.updatePageFormFieldValues(
    writer.getPage(0), dictionary
)

for record in df.to_dict(orient="records"):
  output_path = output_dir / f"{record['insured_name']}" / f"{df['insured_name'].values[0]} - {pdf_filename}"
  output_path.parent.mkdir(exist_ok=True)
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)