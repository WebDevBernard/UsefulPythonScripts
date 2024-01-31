from helper_fn import base_dir
from file_writing_fn import (plumber_draw_rect, plumber_draw_from_pg_and_coords, find_table_dict, get_text_words,
                             write_text_coords, get_text_blocks, search_using_dict, search_with_crop,
                             search_from_pg_num, get_pdf_fieldnames)

# change crop or search here
pg = []  # must be empty list
coords = () # must be tuple to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf_file in pdf_files:
        # get_pdf_fieldnames()
        # Comment to toggle extraction method
        block_dict = get_text_blocks(pdf_file)    # 1 find by text blocks
        # table_dict = find_table_dict(pdf_file)    # 2 find by table
        word_dict = get_text_words(pdf_file)      # 3 find by individual words
        # Assign dicts to write coordinates to file, or preview bbox
        write_text_coords(pdf_file, block_dict, 0, word_dict)  # dicts falsey = off
        # plumber_draw_rect(pdf_file, word_dict, 8, 300)      # field_dict# t falsey = off

