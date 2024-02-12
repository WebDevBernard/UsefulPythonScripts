import fitz
from pathlib import Path
from debug_functions import base_dir
from file_writing_functions import (search_first_page, search_for_matches,
                                    search_for_input_dict, get_broker_copy_pages)
from coordinates import doc_type, keyword, target_dict

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            # Every file requires these inputs:
            print(f"\n<==========================>\n\nFilename is: {Path(pdf).stem}{Path(pdf).suffix} ")

            # Determines type of pdf by scanning page 1 and an area of the page matching a single keyword
            type_of_pdf = search_first_page(doc, doc_type)
            print(f"This is a {type_of_pdf} policy.")

            # Look for a condition to stop search, so it doesn't go into the wordings page
            pg_list = get_broker_copy_pages(doc, type_of_pdf, keyword)
            print(f"\nBroker copies / coverage summary / certificate of property insurance located on pages:\n{pg_list}\n")

            # Extract the dictionary containing all the values I am looking for:
            input_dict = search_for_input_dict(doc, pg_list)
            print(search_for_matches(doc, input_dict, type_of_pdf, target_dict))
            print(f"\n<==========================>\n")

            # need to clean data
            # need to append to pandas df


