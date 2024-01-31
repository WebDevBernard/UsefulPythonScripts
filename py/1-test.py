from helper_fn import base_dir
from file_writing_fn import (plumber_draw_from_pg_and_coords, find_table_dict, get_text_words,
                             get_text_blocks, search_using_dict, search_with_crop, search_from_pg_num)


# change crop or search here
pg = [1]
first_keyword = "Property Coverages"
second_keyword = "LOCATION"
keyword_matching_crop = "Previous"
coords = [(425.3039855957031, 291.7948913574219, 448.932373046875, 299.81011962890625)]  # must be tuple in list to work

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    for pdf_file in pdf_files:
        # Comment to toggle extraction method
        block_dict = get_text_blocks(pdf_file)    # 1 find by text blocks
        table_dict = find_table_dict(pdf_file)    # 2 find by table
        word_dict = get_text_words(pdf_file)       # 3 find by individual words

        list_of_pg_no_matching_keyword = (search_using_dict(block_dict, first_keyword))
        list_of_pg_no_matching_second_keyword = search_from_pg_num(word_dict, list_of_pg_no_matching_keyword, second_keyword)
        list_of_pg_no_matching_crop = (search_with_crop(pdf_file, list_of_pg_no_matching_second_keyword, keyword_matching_crop, coords))

        # Every file requires these inputs:
        #   type of pdf / coordinates of p0 / group of words to search / word that can only appear on that page /
        #       coordinates of page where value appears

        # The first search is scanning pdf[0] and cropping p[0] on an area of the page to get a matching word

        # Once the type of pdf is determined, search through all pages of the pdf continue after finding broker copy
        #       and break after last instance of broker copy

        # Loop through the remaining pages and find group of words that can only appear on the of the value you are
        #       looking for

        # Loop again to find a word match that also appears on the same page

        # using the list of pages with numbered matches, read the list with matching numbers and determine how to much
        #       + index to reach the position alternatively target using set coordinates on pages that have fixed
        #         position like intact and family policies

        print(f"\n  Keyword Matches in Order of: \n")
        print("Using Group of Words:       ", list_of_pg_no_matching_keyword)
        print("Using a Single Word:        ", list_of_pg_no_matching_second_keyword)
        print("Using Crop From Coordinates:", list_of_pg_no_matching_crop)
        plumber_draw_from_pg_and_coords(pdf_file, pg, 0, 300)  # coords falsey = off
