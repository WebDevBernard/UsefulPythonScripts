import fitz
from helper_funtions import newline_to_list, newline_to_space

def find_match_using_pg_num(field_dict, page_number_list, text):
    for key in page_number_list:
        for page in field_dict[key]:
            print(page)

def search_with_crop(pdf, nums_list, text, rects):
    dlist = []
    doc = fitz.open(pdf)
    for key in nums_list:
        page = doc[key - 1]
        wlist = []
        for rect in rects:
            text_line = page.get_text("words", clip=rect)
            wlist.extend(text_line)
        for word in wlist:
            if text in word[4]:
                if key not in dlist:
                    dlist.append(key)
    return dlist

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
    doc = fitz.open(pdf)
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
    doc.close()
    return field_dict

def get_text_blocks(pdf):
    doc = fitz.open(pdf)
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("blocks")
        text_boxes = [inner_list[4] for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
    doc.close()
    return newline_to_list(field_dict)

def get_text_words(pdf):
    doc = fitz.open(pdf)
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text("words")
        text_boxes = [inner_list[4] for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
        # field_dict[page_number + 1] = list(zip(text_boxes, text_coords))
    doc.close()
    return field_dict