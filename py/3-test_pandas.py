import fitz
import os
import pandas as pd
from collections import defaultdict
from pathlib import Path
from debugging import base_dir
from file_writing_functions import (search_first_page, search_for_matches,
                                    search_for_input_dict, get_broker_copy_pages, format_policy, create_pandas_df)
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
            type_of_pdf = search_first_page(doc, doc_type)
            pg_list = get_broker_copy_pages(doc, type_of_pdf, keyword)
            input_dict = search_for_input_dict(doc, pg_list)
            dict_items = search_for_matches(doc, input_dict, type_of_pdf, target_dict)
            cleaned_data = format_policy(dict_items, type_of_pdf)
            print(cleaned_data)
            try:
                df = create_pandas_df(cleaned_data)
            except KeyError:
                continue
            print(f"\n<==========================>\n")

