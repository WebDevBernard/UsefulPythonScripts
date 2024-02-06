import fitz
from pathlib import Path
from helper_fn import base_dir
from file_writing_fn import (search_using_dict, search_with_crop, search_for_pg_limit,
                             search_for_dict, search_for_name_and_address)
from coordinates import doc_type, stop_condition, address_block

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
            type_of_pdf = search_with_crop(doc, [1], doc_type)
            print(f"This is a {type_of_pdf} policy.")

            # Look for a condition to stop search, so it doesn't go into the wordings page
            pg_limit = search_for_pg_limit(doc, type_of_pdf, stop_condition)
            print(f"Search until page {pg_limit}.")

            # Get the text block with the named insured and mailing address
            name_and_address = search_for_name_and_address(doc, type_of_pdf, address_block)
            print(f"\nName and address is:\n{name_and_address}")

            # Extract the list of dictionary containing all the values I am looking for:
            words = search_for_dict(doc, pg_limit)
            # print(f"\nThe {type_of_pdf} Dictionary:\n{words}\n")

            # Find group of words that can match the page of the thing you are looking for
            find_pg_with_keyword = search_using_dict(words, "Location Description")
            print(f"The keyword is found on {find_pg_with_keyword},")

            # Match pg # to the possible postion with list of words and index to position

            # Pages with fixed positions can just get the value from coords without any pattern matching


