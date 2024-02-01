import fitz
from helper_fn import base_dir
from file_writing_fn import (search_using_dict, search_with_crop, search_for_pg_limit,
                             search_for_dict)

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    pdf = list(pdf_files)[0]
    with fitz.open(pdf) as doc:
        # Every file requires these inputs:
        print(f"\n  Keyword Matches in Order of: \n")
        # Determines type of pdf by scanning page 1 and an area of the page matching a single keyword
        doc_type = {
            "Aviva": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
            "Family": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
            "Intact": ["Ageasdasdnt", (22342342345, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
            "Wawanesa": ["Policy", (36.0, 187.6835479736328, 62.479373931884766, 197.7382354736328)],
        }
        keyword_1 = search_with_crop(doc, [1], doc_type)
        print(f"This is a {keyword_1} policy:")

        # Look for a condition to stop search, so it doesn't go into the wordings page
        stop_condition = {
            "Wawanesa": ["STATUTORY CONDITIONS", (242.8795166015625, 99.35482788085938, 371.42803955078125, 110.61607360839844)]
        }
        pg_limit = search_for_pg_limit(doc, keyword_1, stop_condition)
        print(f"search until page {pg_limit},")

        # Extract the list of dictionary containing all the values I am looking for:
        field_dict = search_for_dict(doc, pg_limit)

        # Find group of words that can match the page of the thing you are looking for
        find_pg_with_keyword = search_using_dict(field_dict, "Location Description")

        print(f"the keyword is found on {find_pg_with_keyword},")
        # Loop again to find a word match that also appears on the same page
        print("Using a Single Word:        ", )

        # Match pg # to the possible postion with list of words and index to position
        # Get the text block with the named insured and mailing address



