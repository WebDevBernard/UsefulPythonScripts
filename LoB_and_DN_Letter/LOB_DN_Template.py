# https://github.com/Sven-Bo/word-generator-based-on-excel-list
from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl
from datetime import datetime, timedelta

base_dir = Path(__file__).parent.parent
word_template_path = base_dir / "word.docx"  # name of word doc
excel_path = base_dir / "excel.xlsx"  # name of excel
output_dir = base_dir / datetime.today().strftime("%B %d, %Y %H%M") # name of output folder
output_dir.mkdir(exist_ok=True)

df = pd.read_excel(excel_path, sheet_name="Sheet1")

df["effective_date"] = df["effective_date"].dt.strftime("%B %d, %Y")
thirty_before_effective = pd.to_datetime(df["effective_date"], format="%B %d, %Y") - timedelta(days=30)
df["thirty_before_effective"] = thirty_before_effective.dt.strftime("%B %d, %Y")
df["today"] = datetime.today().strftime("%B %d, %Y")

for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / \
        f"{record['insured_name']}-{record['policy_number']}.docx"
    doc.save(output_path)
