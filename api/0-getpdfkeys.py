import pprint
from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
pdf_folder = base_dir / "input"

for path in Path(pdf_folder).glob("*.pdf"):
    data = PdfReader(path).getFormTextFields()
    fields = [str(x) for x in data]
    print(f"<======{path}======>")
    pprint.pprint(data)

