import fitz
import pdfplumber
from helper_fn import base_dir
from file_writing_fn import (plumber_draw_rect, plumber_draw_from_pg_and_coords, find_table_dict, get_text_words,
                             write_text_coords, get_text_blocks, get_pdf_fieldnames)

# change crop or search here
pg = []  # must be empty list
coords = ()  # must be tuple to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            # GET FILLABLE PDF KEYS
            get_pdf_fieldnames(doc)

            # WRITE TXT COORDS
            block_dict = get_text_blocks(doc)    # 1 find by text blocks
            # table_dict = find_table_dict(doc)    # 2 find by table This is very slow some reason
            word_dict = get_text_words(doc)      # 3 find by individual words
            write_text_coords(doc, block_dict, 0, word_dict)  # dicts falsey = off


        with pdfplumber.open(pdf) as doc:
            # PREVIEW FROM PDF BBOXes
            plumber_draw_rect(doc, 0, 8, 300)      # field_dict# t falsey = off
            plumber_draw_from_pg_and_coords(doc, pg, coords, 300)  # coords falsey = off

