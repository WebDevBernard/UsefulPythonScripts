import fitz
from helper_fn import base_dir
from file_writing_fn import search_with_crop, search_for_name_and_address, search_for_dict
from coordinates import doc_type, address_block
# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            type_of_pdf = search_with_crop(doc, [1], doc_type)
            name_and_address = search_for_name_and_address(doc, type_of_pdf, address_block)
            words = search_for_dict(doc, 1)
            print(f"\nThe {type_of_pdf} Dictionary:\n{words}\n")

