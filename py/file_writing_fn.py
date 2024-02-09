import re
from collections import defaultdict
from helper_fn import base_dir, unique_file_name
from pathlib import Path
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfWriter

# <=================== START OF PYMUPDF FUNCTIONS


# 1st page search to determine type of pdf file
def search_first_page(doc, field_dict):
    page = doc[0]
    for key, field in field_dict.items():
        keyword = field_dict[key][0]
        coords = field_dict[key][1]
        text_block = page.get_text("text", clip=coords)
        if keyword.casefold() in text_block.casefold():
            return key

# 2nd search for address block
def search_for_name_and_address(doc, type_of_pdf, page_coords):
    pg_num = page_coords[type_of_pdf][0]
    coords = page_coords[type_of_pdf][1]
    wlist = doc[pg_num - 1].get_text("blocks", clip=coords)
    text_boxes = [inner_list[4].split("\n") for inner_list in wlist]
    return text_boxes

# 3rd Get the pages with the broker copies
# regex searches is for wawanesa
# set searches is for aviva
# list searches is for intact
# except if no coordinates to search, just loop through all pages

def get_broker_copy_pages(doc, type_of_pdf, keyword):
    pg_list = []
    try:
        for i, pg_num in enumerate(doc):
            kw = keyword[type_of_pdf][0]
            coords = keyword[type_of_pdf][1]
            page = doc[i]
            if isinstance(kw, re.Pattern):
                regex_word = page.get_text("blocks", clip=coords)
                for w in regex_word:
                    if kw.search(w[4]):
                        pg_list.append(i + 1)
            elif isinstance(kw, set):
                stop_word = page.search_for(list(kw)[0], clip=coords)
                if stop_word:
                    for j in range(1, i + 1):
                        pg_list.append(j)
                    break
            elif isinstance(kw, list):
                multi_word = page.search_for(kw[0], clip=coords[0])
                multi_word_2 = page.search_for(kw[1], clip=coords[1])
                if multi_word or multi_word_2:
                    pg_list.append(i + 1)
            else:
                single_word = page.search_for(kw, clip=coords)
                if single_word:
                    pg_list.append(i + 1)
    except KeyError:
        for k in range(doc.page_count):
            pg_list.append(k + 1)
    return pg_list

# 4th search to find the dictionary from the relevant pages
def search_for_wlist(doc, pg_list):
    newlist = []
    for page_num in pg_list:
        page = doc[page_num - 1]
        wlist = page.get_text("blocks")
        for w in wlist:
            newlist.append(w[4].split("\n"))
    return newlist


# 5th find the keys for each matching field and increment outer and inner index of nested list
def search_for_matches(nested_list, type_of_pdf, word_dict):
    field_dict = defaultdict(list)
    try:
        words = word_dict[type_of_pdf]
        for i, sublist in enumerate(nested_list):
            for k, keyword in words.items():
                if any(keyword[0] in s for s in sublist):
                    word = nested_list[i + keyword[1]][keyword[2]]
                    if word not in field_dict[k]:
                        field_dict[k].append(word)
    except KeyError:
        return "Insurer Key does not exist"
    return field_dict


# <=================== END OF PYMUPDF FUNCTIONS

# find blocks words sentences and returns the word and coords as a dictionary of pages containing list of strings and
# their bbox coordinates
def get_text_blocks(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("blocks")
        text_boxes = [inner_list[4].split("\n") for inner_list in wlist]
        text_coords = [inner_list[:4]for inner_list in wlist]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
    return field_dict

# Same as above but finds individual text words (pymupdf debugging)
def get_text_words(doc):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("words")
        text_boxes = [inner_list[4] for inner_list in wlist]
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

# Used for writing coordinates to txt (pymupdf debugging)
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

# Used for filling docx
def write_to_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "input" / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    output_path.parent.mkdir(exist_ok=True)
    doc.save(unique_file_name(output_path))

# Used for fillable pdf's
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

# Opens the default image viewer to see what bbox look like
def plumber_draw_rect(doc, field_dict, pg_limit, dpi):
    if field_dict:
        for page_number in range(len(doc.pages)):
            if page_number + 1 in field_dict and pg_limit >= page_number + 1:
                dict_list = field_dict[page_number + 1]
                doc.pages[page_number].to_image(resolution=dpi).draw_rects([x[1] for x in dict_list]).show()

# Same as above but based on specified coordinates instead
def plumber_draw_from_pg_and_coords(doc, pages, coords, dpi):
    if coords:
        for page_num in pages:
            doc.pages[page_num - 1].to_image(resolution=dpi).draw_rect(coords).show()
