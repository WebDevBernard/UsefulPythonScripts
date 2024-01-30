
from pathlib import Path
from PyPDF2 import PdfReader
import fitz

base_dir = Path(__file__).parent.parent
pdf_folder = base_dir / "input"
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

for pdf in Path(pdf_folder).glob("*.pdf"):

    # print(f"\n{emoji}  PDF Filename:{Path(pdf).stem} {emoji}\n")
    # with open( output_dir / f"{Path(pdf_path).stem} Fillable PDF Fields.txt", 'w') as file:
    #     for i, value in enumerate(data):
    #         file.write(f"#{i}: {value}\n")
    #         print(f"#{i}: {value}")
    doc = fitz.open(pdf)
    for page in doc:
        # print(page.get_links())
        for i, field in enumerate(page.widgets()):
            print(i, field.field_name, field.xref, field.field_value, field.button_states())
        # for annot in page.annots():
        #     print(annot)

