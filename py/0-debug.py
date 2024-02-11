import fitz
import pdfplumber
from helper_fn import base_dir, calculate_target_coords
from file_writing_fn import (plumber_draw_rect, plumber_draw_from_pg_and_coords, find_table_dict, get_text_words,
                             write_text_coords, get_text_blocks, get_pdf_fieldnames)

# change crop or search here
pg = []  # use empty list to toggle off screen preview
input_coords = (141.35000610351562, 507.9750061035156, 174.02598571777344, 525.7430419921875)
target_coords = (30.700000762939453, 526.6749877929688, 545.6491088867188, 536.2930297851562)
direction = "down"
x_y_relative = calculate_target_coords(input_coords, target_coords, direction, False)
debug_relative = calculate_target_coords(input_coords, target_coords, direction, True)
print(f"Coordinates Drawn On: {debug_relative}\nCoordinates to Copy/Paste: {x_y_relative}")
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
            # t = find_table_dict(doc)    # 2 find by table This is very slow some reason
            w = get_text_words(doc)      # 3 find by individual words
            # write_text_coords(pdf, b, 0, w)  # dicts falsey = off

        with pdfplumber.open(pdf) as doc:
            # PREVIEW FROM PDF BBOXes
            # plumber_draw_rect(doc, w, 1, 300)      # field_dict# t falsey = off
            plumber_draw_from_pg_and_coords(doc, pg, coords, 300)  # coords falsey = off

