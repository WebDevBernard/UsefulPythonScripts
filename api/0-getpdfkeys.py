import pprint
from pathlib import Path
from PyPDF2 import PdfReader

pdf_filename = ["GORE - Rented Questionnaire.pdf",
                "LOB - Family Blank.pdf",
                "Optimum West Rental Q.pdf",
                "Wawa Personal Information and Credit Consent Form 8871.pdf",
                "WAWA Rental Condo Questionnaire.pdf",
                "wawa rented dwelling Q.pdf",
                "Binder Fee Invoice - Cedar.pdf",
                "Wawa monthly payplan Authorization Form, Form 8003GIS.pdf"
                ]
base_dir = Path(__file__).parent.parent


def readPdf(pdf):
    pdf_path = base_dir / "input" / pdf
    return pdf_path


for pdf in pdf_filename:
    pdf_path = readPdf(pdf)
    reader = PdfReader(pdf_path)
    fields = [str(x) for x in reader.getFormTextFields()]
    print("<========" + pdf + "========>")
    pprint.pprint(fields)
