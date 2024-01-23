from pathlib import Path
from PyPDF2 import PdfReader

base_dir = Path(__file__).parent.parent
pdf_folder = base_dir / "templates"
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

for pdf_path in Path(pdf_folder).glob("*.pdf"):
    data = PdfReader(pdf_path).get_form_text_fields()
    print(f"\n<======PDF Filename:{Path(pdf_path).stem}======>\n")
    with open( output_dir / f"{Path(pdf_path).stem} Fillable PDF Fields.txt", 'w') as file:
        for i, value in enumerate(data):
            file.write(f"#{i}: {value}\n")
            print(f"#{i}: {value}")
