import fitz
import os
import pandas as pd
from collections import defaultdict
from pathlib import Path
from debug_functions import base_dir
from file_writing_functions import (search_first_page, search_for_matches, format_dict_items, write_to_new_docx,
                                    search_for_input_dict, get_broker_copy_pages, create_pandas_df)
from coordinates import doc_type, keyword, target_dict, dict_of_keywords

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")


def main():

    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            # Every file requires these inputs:
            print(f"\n<==========================>\n\nFilename is: {Path(pdf).stem}{Path(pdf).suffix} ")

            # Determines type of pdf by scanning page 1 and an area of the page matching a single keyword
            type_of_pdf = search_first_page(doc, doc_type)

            # Look for a condition to stop search, so it doesn't go into the wordings page
            pg_list = get_broker_copy_pages(doc, type_of_pdf, keyword)

            # Extract the dictionary containing all the values I am looking for:
            input_dict = search_for_input_dict(doc, pg_list)
            dict_items = search_for_matches(doc, input_dict, type_of_pdf, target_dict)
            # Clean data
            cleaned_data = format_dict_items(dict_items)
            print(cleaned_data)
            # format_dict_items(dict_items,type_of_pdf,dict_of_keywords)
            # Append to Pandas DF
            # df = create_pandas_df(cleaned_data)
            # print(df)
            # print(f"\n<==========================>\n")


if __name__ == "__main__":
   main()

