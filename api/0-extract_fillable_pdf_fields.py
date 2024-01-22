from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
pdf_folder = base_dir / "templates"

for path in Path(pdf_folder).glob("*.pdf"):
    data = PdfReader(path).getFormTextFields()
    print("\n")
    print(f"<======PDF Filename:{path}======>")
    print("\n")
    for value in data:
        print(f"{value}")

