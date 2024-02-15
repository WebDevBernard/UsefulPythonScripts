import fitz
import pdfplumber
from debugging import (base_dir, calculate_target_coords, plumber_draw_rect, plumber_draw_from_pg_and_coords,
                       find_table_dict, get_text_words, write_text_coords, get_text_blocks, get_pdf_fieldnames)

# change crop or search here
pg = []  # use empty list to toggle off-screen preview
loc_coords = 1

input_coords = (32.880001068115234, 757.760009765625, 552.3679809570312, 765.7599487304688)
target_coords = (300.0, 767.9199829101562, 349.54901123046875, 774.9200439453125)
direction = "down"
debug_relative = calculate_target_coords(input_coords, target_coords, direction)

if loc_coords:
    coords = loc_coords  # must be tuple to work
else:
    coords = debug_relative

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            get_pdf_fieldnames(doc)

            b = get_text_blocks(doc)    # 1 find by text blocks
            # t = find_table_dict(doc)    # 2 find by table
            w = get_text_words(doc)      # 3 find by individual words
            write_text_coords(pdf, b, 0, w)  # dicts falsey = off

        with pdfplumber.open(pdf) as doc:
            # plumber_draw_rect(doc, b, 0, 300)      # field_dict# falsey = off
            plumber_draw_from_pg_and_coords(doc, pg, coords, 300)  # coords falsey = off

