from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
# pdf_path = base_dir / "input" / "GORE - Rented Questionnaire.pdf"
# pdf_path = base_dir / "input" / "LOB - Family Blank.pdf"
# pdf_path = base_dir / "input" / "Optimum West Rental Q.pdf"
pdf_path = base_dir / "input" / "Wawa Personal Information and Credit Consent Form 8871.pdf"
# pdf_path = base_dir / "input" / "WAWA Rental Condo Questionnaire.pdf"
# pdf_path = base_dir / "input" / "wawa rented dwelling Q.pdf"
reader = PdfReader(pdf_path)
fields = [str(x) for x in reader.getFields().keys()]
print(fields)