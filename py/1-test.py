import fitz
from pathlib import Path
from helper_fn import base_dir
from file_writing_fn import (search_first_page, search_for_matches,
                             search_for_wlist, search_for_name_and_address, get_broker_copy_pages)
from coordinates import doc_type, keyword, address_block, targets

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            # Every file requires these inputs:
            print(f"\nFilename is: {Path(pdf).stem}{Path(pdf).suffix} ")

            # Determines type of pdf by scanning page 1 and an area of the page matching a single keyword
            type_of_pdf = search_first_page(doc, doc_type)
            print(f"This is a {type_of_pdf} policy.")

            # Get the text block with the named insured and mailing address
            name_and_address = search_for_name_and_address(doc, type_of_pdf, address_block)
            print(f"\nName and address is:\n{name_and_address}")

            # Look for a condition to stop search, so it doesn't go into the wordings page
            pg_list = get_broker_copy_pages(doc, type_of_pdf, keyword)
            print(f"\nBroker copies / coverage summary located on pages:\n{pg_list}")

            # Extract the list of dictionary containing all the values I am looking for:
            words = search_for_wlist(doc, pg_list)
            # print(f"\nThe {type_of_pdf} Dictionary:\n{words}\n")


            # Find group of words that can match the page of the thing you are looking for
            print(search_for_matches(words, type_of_pdf, targets))
            # print(f"The keyword is found on {find_pg_with_keyword},")

            # Match pg # to the possible postion with list of words and index to position

            # Pages with fixed positions can just get the value from coords without any pattern matching


