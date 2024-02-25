import fitz
import pdfplumber
from pathlib import Path

base_dir = Path(__file__).parent.parent
# change crop or search here
pg = [1]  # use empty list to toggle off-screen preview
coords = (250, 764.2749633789062, 360, 773.8930053710938)

# find blocks words sentences and returns the word and coords as a dictionary of pages containing list of strings and
# their bbox coordinates
def get_text_blocks(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("blocks")
        text_boxes = [list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
    return field_dict


# Same as above but finds individual text words (pymupdf debugging)
def get_text_words(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("words")
        text_boxes = [list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
    return field_dict


# find words based on table looking things (pymupdf debugging)
def find_table_dict(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        tlist = []
        row_coords = []
        find_tables = page.find_tables()
        for table in find_tables:
            tlist.extend(table.extract())
            row_coords.append(table.bbox)
        field_dict[page_number + 1] = field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                                                     zip(tlist, row_coords)]
    return field_dict


# Gets pdf key for fillable pdfs (pymupdf debugging)
def get_pdf_fieldnames(doc):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    for page in doc:
        for i, field in enumerate(page.widgets()):
            print(i, field.field_name, field.xref, field.field_value)


# Opens the default image viewer to see what bbox look like
def plumber_draw_rect(doc, field_dict, pg_limit, dpi):
    if field_dict:
        for page_number in range(len(doc.pages)):
            if page_number + 1 in field_dict and pg_limit >= page_number + 1:
                dict_list = field_dict[page_number + 1]
                doc.pages[page_number].to_image(resolution=dpi).draw_rects([x[1] for x in dict_list]).show()


# Same as above but based on specified coordinates instead
def plumber_draw_from_pg_and_coords(doc, pages, coords, dpi):
    for page_num in pages:
        doc.pages[page_num - 1].to_image(resolution=dpi).draw_rect(coords).show()


# prints coordinates using input and target
def calculate_target_coords(input_coords, target_coords, direction):
    if direction == "down":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_x_0 = x0_1 + x0_2 - x0_1
        debug_result_x_1 = x1_1 + x1_2 - x1_1
        debug_result_y_0 = y0_1 + y0_2 - y0_1
        debug_result_y_1 = y1_1 + y1_2 - y1_1
        result_x_0 = x0_2 - x0_1
        result_x_1 = x1_2 - x1_1
        result_y_0 = y0_2 - y0_1
        result_y_1 = y1_2 - y1_1
        x_y_relative = (result_x_0, result_y_0, result_x_1, result_y_1)
        debug_relative = (debug_result_x_0, debug_result_y_0, debug_result_x_1, debug_result_y_1)
        print(f"Coordinates Drawn On: {debug_relative}\nCoordinates to Copy/Paste: {x_y_relative}")
        return debug_relative

    if direction == "right":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_x_0 = x0_1 + x0_2 - x0_1
        debug_result_x_1 = x1_1 + x1_2 - x1_1
        result_x_0 = x0_2 - x0_1
        result_x_1 = x1_2 - x1_1
        x_y_relative = (result_x_0, 0, result_x_1, 0)
        debug_relative = (debug_result_x_0, y0_1, debug_result_x_1, y1_1)
        print(f"Coordinates Drawn On: {debug_relative}\nCoordinates to Copy/Paste: {x_y_relative}")
        return debug_relative


def write_txt_to_file_dir(dir_path, field_dict):
    with open(dir_path, 'w', encoding="utf-8") as file:
        for page, value in field_dict.items():
            file.write(f"\n<========= Page: {page} =========>\n")
            for text, box in enumerate(value):
                file.write(f"#{text} : {box}\n")


def write_text_coords(file_name, block_dict, table_dict, word_dict):
    output_dir = base_dir / "output" / Path(file_name).stem
    output_dir.mkdir(exist_ok=True)
    block_dir_path = output_dir / f"block_coordinates_{Path(file_name).stem}.txt"
    table_dir_path = output_dir / f"table_coordinates_{Path(file_name).stem}.txt"
    word_dir_path = output_dir / f"word_coordinates_{Path(file_name).stem}.txt"
    if block_dict:
        write_txt_to_file_dir(block_dir_path, block_dict)
    if table_dict:
        write_txt_to_file_dir(table_dir_path, table_dict)
    if word_dict:
        write_txt_to_file_dir(word_dir_path, word_dict)

def main():
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            get_pdf_fieldnames(doc)

            b = get_text_blocks(doc)    # 1 find by text blocks
            # t = find_table_dict(doc)    # 2 find by table
            w = get_text_words(doc)      # 3 find by individual words
            write_text_coords(pdf, b, 0, w)  # dicts falsey = off

        with pdfplumber.open(pdf) as doc:
            # plumber_draw_rect(doc, b, 8, 300)      # field_dict# falsey = off
            plumber_draw_from_pg_and_coords(doc, pg, coords, 300)  # coords falsey = off


# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    main()
