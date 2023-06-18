import pandas as pd  # pip install pandas openpyxl
from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl
from datetime import datetime, timedelta

#Directory Paths for each file
docx_filename = "Rented Dwelling Quest INTACT.docx" #filename
base_dir = Path(__file__).parent.parent
LOB_template_path = base_dir / "input" / docx_filename  # name of LOB doc
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

# Loop generates LOB and DN
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(LOB_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['insured_name']}" / f"{record['insured_name']} - {docx_filename}"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(output_path)
