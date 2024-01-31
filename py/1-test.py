from helper_fn import base_dir
from file_writing_fn import (plumber_draw_from_pg_and_coords, find_table_dict, get_text_words,
                             get_text_blocks, search_using_dict, search_with_crop, search_from_pg_num, search_pg_num_stop)
import fitz

# change crop or search here
pg = [8]
coords = (36.0, 769.5787963867188, 576.001220703125, 778.6453857421875)  # must be tuple in list to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    pdf = list(pdf_files)[0]

    # Every file requires these inputs:
    print(f"\n  Keyword Matches in Order of: \n")
    # Determines type of pdf by scanning page 1 and an area of the page matching a single keyword
    doc_type = {
        "Aviva": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
        "Family": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
        "Intact": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
        "Wawanesa": ["Policy", (36.0, 187.6835479736328, 62.479373931884766, 197.7382354736328)],
    }
    keyword_1 = search_with_crop(pdf, [1], doc_type)
    print("Using Crop From Coordinates:", keyword_1)
    # Look for a condition to stop search, so it doesn't go into the wordings page
    stop_condition = {
        "Wawanesa": ["STATUTORY CONDITIONS", (242.8795166015625, 99.35482788085938, 371.42803955078125, 110.61607360839844)]
    }

    print("Finding Pg# to Stop:", search_pg_num_stop(pdf, keyword_1, stop_condition))
    print("Using Group of Words:       ", )
    # Loop again to find a word match that also appears on the same page
    print("Using a Single Word:        ", )
    # using the list of pages with numbered matches, read the list with matching numbers and determine how to much
    #       + index to reach the position alternatively target using set coordinates on pages that have fixed
    #         position like intact and family policies


    # plumber_draw_from_pg_and_coords(pdf, pg, coords, 300)  # coords falsey = off
