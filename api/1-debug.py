from helper_funtions import unique_file_name, remove_newlines, emoji
from pdfplumber_functions import plumber_draw_rect, plumber_extract_text, plumber_draw_from_pg_and_coords
from pymupdf_functions import find_table_dict, get_text_dict
from pathlib import Path
import os

search = "Insured's name and address"
crop = None

base_dir = Path(__file__).parent.parent
pdf_directory = Path(base_dir / "input")
pdf_files = pdf_directory.glob("*.pdf")
pdf_file_paths = []
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)

# Loop through each PDF file and append the full path to the list
for pdf_file in pdf_files:
    file_path = str(pdf_file)
    pdf_file_paths.append(file_path)

for pdf_file in pdf_file_paths:

    # Comment to toggle extraction method
    # field_dict = plumber_extract_text(pdf_file) # find text and bbox of each word looking thing
    field_dict = get_text_dict(pdf_file, crop) # find text line and bbox of each word looking thing
    # field_dict = find_table_dict(pdf_file) # find table and bbox of each word looking thing

    with open(output_dir / f"{Path(pdf_file).stem}.txt", 'w', encoding="utf-8") as file:
        for page, value in field_dict.items():
            file.write(f"\n<========= Page: {page} =========>\n")
            print(f"\n{emoji}   Page: {page} {emoji}\n")
            for text, box in enumerate(value):
                file.write(f"#{text} : {box}\n")
                print(f"#{text} :{box}\n")

        # Comment to toggle image debugger and file previewer
        os.startfile(str(Path(file.name)))
        plumber_draw_rect(pdf_file, field_dict)
        # plumber_draw_from_pg_and_coords()
