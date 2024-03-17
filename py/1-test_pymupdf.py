import fitz
from pathlib import Path
from file_completion_tool import (get_doc_types, get_content_pages, search_for_input_dict, search_for_matches,
                                  format_policy, sort_renewal_list, create_pandas_df, rename_icbc)
from helpers import target_dict, ff

def main():
    for pdf_file in pdf_files:
        with fitz.open(pdf_file) as doc:
            print(Path(pdf_file).stem)
            print(f"\n<==========================>\n\nFilename is: {Path(pdf_file).stem}{Path(pdf_file).suffix} ")
            #
            # # Look through Pdf Pages to find the matching pdf
            # doc_type = get_doc_types(doc)
            # print(f"This is a {doc_type} policy.")
            #
            # # Look for pages with the content for each field
            # pg_list = get_content_pages(doc, doc_type)
            # print(
            #     f"\nBroker copies / coverage summary / certificate of property insurance located on pages:\n{pg_list}\n")
            #
            # # Extract the dictionary containing for all the strings and corresponding coordinates:
            # input_dict = search_for_input_dict(doc, pg_list)
            # dict_items = search_for_matches(doc, input_dict, doc_type, target_dict)
            # print(ff(dict_items[doc_type]))
            # formatted_dict = format_policy(ff(dict_items[doc_type]), doc_type)
            # # print(formatted_dict)
            # # print(create_pandas_df(formatted_dict))
            # # rename_icbc()
            # print(f"\n<==========================>\n")


base_dir = Path(__file__).parent.parent
# input_dir = base_dir / "input"
input_dir = Path.home() / 'Downloads'
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    main()
