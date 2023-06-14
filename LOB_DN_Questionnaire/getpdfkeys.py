from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / "WAWA Rental Condo Questionnaire.pdf"
reader = PdfReader(pdf_path)
fields = [str(x) for x in reader.getFields().keys()]
print(fields)