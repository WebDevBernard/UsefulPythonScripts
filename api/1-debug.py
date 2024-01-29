from helper_functions import base_dir, emoji, unique_file_name, write_text_coords
from pdfplumber_functions import plumber_draw_rect, plumber_draw_from_pg_and_coords
from pymupdf_functions import find_table_dict, get_text_words, get_text_blocks, search_using_dict, search_with_crop, search_from_pg_num

# change crop or search here
pg = [1]
coords = [] # must be tuple in list to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")
pdf_file_paths = []

if __name__ == "__main__":
    for pdf_file in pdf_files:
        file_path = str(pdf_file)
        pdf_file_paths.append(file_path)

    for pdf_file in pdf_file_paths:
        # Comment to toggle extraction method
        block_dict = get_text_blocks(pdf_file)    # 1 find by text blocks
        table_dict = find_table_dict(pdf_file)    # 2 find by table
        word_dict = get_text_words(pdf_file)      # 3 find by individual words
        # Assign dicts to write coordinates to file, or preview bbox
        write_text_coords(pdf_file, block_dict, 0, word_dict)  # dicts falsey = off
        plumber_draw_rect(pdf_file, 0, 8, 300)      # field_dict falsey = off
        plumber_draw_from_pg_and_coords(pdf_file, pg, coords, 300)  # coords falsey = off
