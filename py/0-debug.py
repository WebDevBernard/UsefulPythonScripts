import fitz
import pdfplumber
from debug_functions import (base_dir, calculate_target_coords, plumber_draw_rect, plumber_draw_from_pg_and_coords,
                             find_table_dict, get_text_words, write_text_coords, get_text_blocks, get_pdf_fieldnames)

# change crop or search here
pg = []  # use empty list to toggle off-screen preview
input_coords = (270.239990234375, 582.1600341796875, 544.447998046875, 592.1600341796875)
target_coords = (46.31999969482422, 629.3600463867188, 512.7559814453125, 642.8800048828125)
direction = "down"
debug_relative = calculate_target_coords(input_coords, target_coords, direction)

coords = debug_relative  # must be tuple to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            # GET FILLABLE PDF KEYS
            # get_pdf_fieldnames(doc)

            # WRITE TXT COORDS
            b = get_text_blocks(doc)    # 1 find by text blocks
            # t = find_table_dict(doc)    # 2 find by table
            w = get_text_words(doc)      # 3 find by individual words
            write_text_coords(pdf, b, 0, w)  # dicts falsey = off

        with pdfplumber.open(pdf) as doc:
            # PREVIEW FROM PDF BBOXes
            # plumber_draw_rect(doc, w, 1, 300)      # field_dict# falsey = off
            plumber_draw_from_pg_and_coords(doc, pg, coords, 300)  # coords falsey = off

