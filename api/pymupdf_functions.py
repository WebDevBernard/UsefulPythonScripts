import fitz

def find_table_dict(pdf_path):
    doc = fitz.open(pdf_path)
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        text_boxes = []
        row_coords = []
        find_tables = page.find_tables()
        for table in find_tables:
            text_boxes.extend(table.extract())
            row_coords.append(table.bbox)
        field_dict[page_number + 1] = field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                                                 zip(text_boxes, row_coords)]
    doc.close()
    return field_dict

def get_text_dict(pdf_path, rect):
    doc = fitz.open(pdf_path)
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        text_lines = page.get_text("blocks", clip=rect)
        text_boxes = [inner_list[4] for inner_list in text_lines]
        text_coords = [inner_list[:4] for inner_list in text_lines]
        field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in
                                       zip(text_boxes, text_coords)]
    doc.close()
    return field_dict