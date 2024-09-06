from pathlib import Path

import fitz
import pdfplumber
from tabulate import tabulate

draw_rect_on_page_num = [1]
draw_rect_from_coords = (35.49599838256836,
                153.38099670410156,
                150,
                228.67138671875)
input_coords = (284.2560119628906, 35.721092224121094, 329.5166015625, 44.79716873168945)
# set target coords to (+, -, -, -) to get the reverse
target_coords = (
    283.32000732421875, 46.81807327270508, 337.2813720703125, 56.24658203125
)
turn_on_draw_rect_all = []
turn_on_pdf_field_names = []
turn_on_write_text_coords = [1]


# Gets pdf key for fillable pdfs (pymupdf debugging)
def get_pdf_fieldnames(filename, doc):
    if turn_on_pdf_field_names:
        output_dir = base_dir / "output" / Path(filename).stem
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"(FieldNames) {Path(filename).stem}.txt"
        with open(output_path, "w", encoding="utf-8") as file:
            for index, page in enumerate(doc):
                file.write(f"\nPage: {index + 1}\n")
                print(f"\nPage {index + 1}\n")
                for index1, widget in enumerate(page.widgets()):
                    file.write(f"{widget.field_name}")
                    print(f"{index1}{" " * 9}{widget.field_name}")


# def get_pdf_annot(doc):
#     for page_number in range(doc.page_count):
#         page = doc[page_number]
#         page.clean_contents()
#         xref = page.get_contents()[0]
#         cont = bytearray(page.read_contents())
#
#         print(cont)
#         while True:
#             i1 = cont.find(b"/Artifact")  # start of definition
#             if i1 < 0:
#                 break  # none more left: done
#             i2 = cont.find(b"EMC", i1)  # end of definition
#             cont[i1 - 2 : i2 + 3] = b""  # remove the full definition source "q ... EMC"
#         doc.update_stream(xref, cont)  # replace the original source
#         doc.ez_save(Path.home() / "Desktop" / "x.pdf")


# prints coordinates using input and target
def calculate_target_coords(input_coords, target_coords):
    if input_coords and target_coords:
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
        debug_relative = (
            debug_result_x_0,
            debug_result_y_0,
            debug_result_x_1,
            debug_result_y_1,
        )
        print(
            f"Coordinates Drawn On: {debug_relative}\nCoordinates to Copy/Paste: {x_y_relative}"
        )
        return debug_relative


# Get list of word looking things (pymupdf debugging)
def get_text_blocks(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("blocks")
        text_boxes = [
            list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist
        ]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [
            [elem1, elem2] for elem1, elem2 in zip(text_boxes, text_coords)
        ]
    return field_dict


# Same as above but finds individual text words (pymupdf debugging)
def get_text_words(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("words")
        text_boxes = [
            list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist
        ]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [
            [elem1, elem2] for elem1, elem2 in zip(text_boxes, text_coords)
        ]
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
        field_dict[page_number + 1] = field_dict[page_number + 1] = [
            [elem1, elem2] for elem1, elem2 in zip(tlist, row_coords)
        ]
    return field_dict


# Opens the default image viewer to see what bbox look like
def plumber_draw_rect(doc, field_dict, dpi):
    if turn_on_draw_rect_all:
        for page_number in range(len(doc.pages)):
            dict_list = field_dict[page_number + 1]
            doc.pages[page_number].to_image(resolution=dpi).draw_rects(
                [x[1] for x in dict_list]
            ).show()


# Same as above but based on specified coordinates instead
def plumber_draw_from_pg_and_coords(doc, pages, coords1, dpi):
    for page_num in pages:
        print(doc.pages[page_num - 1].crop(coords1).extract_text())
        doc.pages[page_num - 1].to_image(resolution=dpi).draw_rect(coords1).show()


def write_txt_to_file_dir(dir_path, field_dict):
    with open(dir_path, "w", encoding="utf-8") as file:
        for page, data in field_dict.items():
            file.write(f"Page: {page}\n")
            try:
                file.write(
                    f"{tabulate(data, ["Keywords", "Coordinates"], tablefmt="grid", maxcolwidths=[
                    50, None])}\n"
                )
            except IndexError:
                continue


def write_text_coords(file_name, block_dict, table_dict, word_dict):
    if turn_on_write_text_coords:
        output_dir = base_dir / "output" / Path(file_name).stem
        output_dir.mkdir(parents=True, exist_ok=True)
        block_dir_path = output_dir / f"(Block Coordinates) {Path(file_name).stem}.txt"
        table_dir_path = output_dir / f"(Table Coordinates) {Path(file_name).stem}.txt"
        word_dir_path = output_dir / f"(Word Coordinates) {Path(file_name).stem}.txt"
        if block_dict:
            write_txt_to_file_dir(block_dir_path, block_dict)
        if table_dict:
            write_txt_to_file_dir(table_dir_path, table_dict)
        if word_dict:
            write_txt_to_file_dir(word_dir_path, word_dict)


def main():
    if not Path(input_dir).exists():
        input_dir.mkdir(exist_ok=True)
        return
    else:
        calculate_target_coords(input_coords, target_coords)
        for pdf in pdf_files:
            with fitz.open(pdf) as doc:
                get_pdf_fieldnames(pdf, doc)

                b = get_text_blocks(doc)  # 1 find by text blocks
                # t = find_table_dict(doc)  # 2 find by table
                w = get_text_words(doc)  # 3 find by individual words
                write_text_coords(pdf, b, 0, w)


            with pdfplumber.open(pdf) as doc:
                plumber_draw_rect(doc, b, 300)
                plumber_draw_from_pg_and_coords(
                    doc, draw_rect_on_page_num, draw_rect_from_coords, 300
                )


base_dir = Path(__file__).parent.parent

# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    main()
