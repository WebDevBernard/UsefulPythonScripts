import pandas as pd  # pip install pandas openpyxl
from datetime import datetime, timedelta
from pathlib import Path
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject

pdf_filename = "Wawa Personal Information and Credit Consent Form 8871.pdf" #filename
base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / pdf_filename
excel_path = base_dir / "input.xlsx"  # name of excel
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

#Pandas reading excel file
df = pd.read_excel(excel_path, sheet_name="Sheet2")

# Formats dates to MMM DD, YYYY
df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")
df["today"] = datetime.today().strftime("%B %d, %Y")

reader = PdfFileReader(pdf_path)
writer = PdfFileWriter()

for pageNum in range(reader.numPages):
  page = reader.getPage(pageNum)
  writer.add_page(page)

for rows in df.to_dict(orient="records"):
  dictionary = {"Policy  Submission Numbers": rows["policy_number"],
              "Insureds Name": rows["insured_name"],
              }
  writer.updatePageFormFieldValues(
    writer.getPage(0), dictionary
  )
  output_path = output_dir / f"{rows['insured_name']} - {pdf_filename}"
  output_path.parent.mkdir(exist_ok=True)
  with open(output_path, "wb") as output_stream:
    writer.write(output_stream)
