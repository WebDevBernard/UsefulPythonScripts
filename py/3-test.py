import fitz
import pikepdf
import os
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

base_dir = Path(__file__).parent.parent
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")
input_path = Path.home() / "Desktop" / "a.pdf"
output_path = Path.home() / "Desktop" / "x.pdf"


def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


def remove_encryption_from_pdf():
    with open(input_path, "rb") as file:
        reader = PdfReader(file)
        if reader.is_encrypted:
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(output_path, "wb") as output_pdf:
                writer.write(output_pdf)

font_color=tuple(value/255 for value in (0, 137, 210))

for pdf in pdf_files:
    # remove_encryption_from_pdf()
    with fitz.open(pdf) as doc:
        for page in doc:
            for index, field in enumerate(page.widgets()):
                if field.field_type == 7:
                    field.field_value = "{0}".format(index)
                    field.update()
                if field.field_type == 2 or field.field_type == 5:
                    field.field_value = True
                    field.update()
                    page.insert_text(field.rect.tl, "{0}".format(index), fontsize=6, color=font_color)
        doc.save(unique_file_name(output_path), garbage=4)

