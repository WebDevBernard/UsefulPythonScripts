import pdfplumber
import fitz
from helper_fn import base_dir, newline_to_list, unique_file_name
from pathlib import Path
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfWriter


# Pymupdf functions
def find_match_using_pg_num(field_dict, page_number_list, text):
    for key in page_number_list:
        for page in field_dict[key]:
            print(page)


def search_with_crop(pdf, pg_num_list, field_dict):
    with fitz.open(pdf) as doc:
        for pg_num in pg_num_list:
            wlist = []
            page_0_index = doc[pg_num - 1]
            for key, field in field_dict.items():
                text_line = page_0_index.get_text("words", field[1])
                wlist.extend(text_line)
                for word in wlist:
                    if field[0].casefold() in word[4].casefold():
                        return key

def search_pg_num_stop(pdf, dict_key, field_dict):
    pg_count = 1
    with fitz.open(pdf) as doc:
        for i, pg_num in enumerate(doc):
            counter = 0
            page_0_index = doc[i]
            word = (field_dict[dict_key][0])
            wlist = page_0_index.get_text("words", field_dict[dict_key][1])
            print(wlist)
            if word in word:
                counter = counter + 1
                if counter >= 1:
                    pg_count += i
                    break
    print(pg_count)
    return pg_count

def search_from_pg_num(field_dict, list_of_nums, text):
    list = []
    for key in list_of_nums:
        for word in field_dict[key]:
            if text.casefold() in word[0].casefold():
                if key not in list:
                    list.append(key)
    return list


def search_using_dict(field_dict, text):
    list = []
    for key, nested_list in field_dict.items():
        for item in nested_list:
            for word in item[0]:
                if text.casefold() in word.casefold():
                    if key not in list:
                        list.append(key)
    return list


def find_table_dict(pdf):
    with fitz.open(pdf) as doc:
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


def get_text_blocks(pdf):
    with fitz.open(pdf) as doc:
        field_dict = {}
        for page_number in range(doc.page_count):
            page = doc[page_number]
            wlist = page.get_text("blocks")
            text_boxes = [inner_list[4] for inner_list in wlist]
            text_coords = [inner_list[:4] for inner_list in wlist]
            field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                           zip(text_boxes, text_coords)]
        return newline_to_list(field_dict)


def get_text_words(pdf):
    with fitz.open(pdf) as doc:
        field_dict = {}
        for page_number in range(doc.page_count):
            page = doc[page_number]
            wlist = page.get_text("words")
            text_boxes = [inner_list[4] for inner_list in wlist]
            text_coords = [inner_list[:4] for inner_list in wlist]
            field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                           zip(text_boxes, text_coords)]
        return field_dict


def get_pdf_fieldnames():
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    for pdf in Path(input_dir).glob("*.pdf"):
        with fitz.open(pdf) as doc:
            for page in doc:
                for i, field in enumerate(page.widgets()):
                    print(i, field.field_name, field.xref, field.field_value)


def write_text_coords(file_name, block_dict, table_dict, word_dict):
    output_dir = base_dir / "output" / Path(file_name).stem
    output_dir.mkdir(exist_ok=True)
    block_dir_path = output_dir / f"block_coordinates_{Path(file_name).stem}.txt"
    table_dir_path = output_dir / f"table_coordinates_{Path(file_name).stem}.txt"
    word_dir_path = output_dir / f"word_coordinates_{Path(file_name).stem}.txt"
    if block_dict:
        with open(block_dir_path, 'w', encoding="utf-8") as file:
            for page, value in block_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")
    if table_dict:
        with open(table_dir_path, 'w', encoding="utf-8") as file:
            for page, value in table_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")
    if word_dict:
        with open(word_dir_path, 'w', encoding="utf-8") as file:
            for page, value in word_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")


def write_to_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "input" / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(unique_file_name(output_path))


def write_to_pdf(pdf, dictionary, rows):
    pdf_path = (base_dir / "input" / "templates" / pdf)
    output_path = base_dir / "output" / f"{rows["named_insured"]} {rows["type"].title()}.pdf"
    output_path.parent.mkdir(exist_ok=True)
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    with open(unique_file_name(output_path), "wb") as output_stream:
        writer.write(output_stream)


def plumber_draw_rect(pdf, field_dict, pg_limit, dpi):
    if field_dict:
        with pdfplumber.open(pdf) as pdf:
            for page_number in range(len(pdf.pages)):
                if page_number + 1 in field_dict and pg_limit >= page_number + 1:
                    dict_list = field_dict[page_number + 1]
                    pdf.pages[page_number].to_image(resolution=dpi).draw_rects([x[1] for x in dict_list]).show()


def plumber_draw_from_pg_and_coords(pdf, pages, coords, dpi):
    if coords:
        with pdfplumber.open(pdf) as pdf:
            for page_num in pages:
                pdf.pages[page_num - 1].to_image(resolution=dpi).draw_rect(coords).show()
