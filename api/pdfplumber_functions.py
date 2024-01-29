import pdfplumber

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
                pdf.pages[page_num - 1].to_image(resolution=dpi).draw_rects([x for x in coords]).show()

def plumber_search_text(pdf, keyword):
    field_dict = {}
    with pdfplumber.open(pdf) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            text_boxes = []
            description = page.search(keyword, case=False)
            for line in description:
                bounding_box = (line['x0'], line['top'], line['x1'], line['bottom'])
                text_boxes.append([line['text'], bounding_box])
            field_dict[page_number + 1] = text_boxes
    return field_dict

def plumber_extract_text(pdf):
    field_dict = {}
    with pdfplumber.open(pdf) as pdf:
        for page_number in range(len(pdf.pages)):
            text_boxes = []
            page = pdf.pages[page_number]
            text_lines = page.extract_words(use_text_flow=True)
            for line in text_lines:
                bounding_box = (line['x0'], line['top'], line['x1'], line['bottom'])
                text_boxes.append([line['text'], bounding_box])
            field_dict[page_number + 1] = text_boxes
    return field_dict

def plumber_extract_tables(pdf_path, ts, bbox):
    field_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            page = page.crop(bbox=bbox)
            row_coords = []
            find_tables = page.find_tables(ts)
            for table in find_tables:
                for row in table.rows:
                    row_coords.append(row.bbox)
            extract_tables = page.extract_tables(ts)
            table_texts = []
            for table in extract_tables:
                table_texts.extend(table)
            field_dict[page_number + 1] = field_dict[page_number + 1] = [[elem1, elem2] for elem1, elem2 in zip(table_texts, row_coords)]
    return field_dict