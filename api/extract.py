from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
pdf_path = base_dir / "input" / "Review Application.pdf"

reader = PdfReader(pdf_path)
page = reader.pages[0]

parts = []

def visitor_body(text, cm, tm, fontDict, fontSize):
    x = tm[4]
    y = tm[5]
    if (y > 3800 and y < 3900):
        parts.append(text)

page.extract_text(visitor_text=visitor_body)
text_body = "".join(parts)

print(text_body)


