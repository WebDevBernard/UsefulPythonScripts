import re
import pandas as pd
from collections import defaultdict
from debug_functions import base_dir, unique_file_name
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfWriter


# 1st page search to determine type of pdf file
def search_first_page(doc, field_dict):
    page = doc[0]
    for key, field in field_dict.items():
        keyword = field_dict[key][0]
        coords = field_dict[key][1]
        text_block = page.get_text("text", clip=coords)
        if keyword.casefold() in text_block.casefold():
            return key


# 2nd Get the pages with the broker copies
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


# 3rd search to find the dictionary from the relevant pages
def search_for_input_dict(doc, pg_list):
    field_dict = {}
    for page_num in pg_list:
        page = doc[page_num - 1]
        wlist = page.get_text("blocks")
        text_boxes = [list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_num] = [[elem1, elem2] for elem1, elem2 in
                                zip(text_boxes, text_coords)]
    return field_dict


# 4th find the keys for each matching field and increment outer and inner index of nested list
def append_word_to_dict(wlist, dict, no_replace):
    for words in wlist:
        word = words[4].strip().split("\n")
        if no_replace:
            dict.append(word)
        if word and word not in dict:
            dict.append(word)


def search_for_matches(doc, input_dict, type_of_pdf, target_dict):
    field_dict = defaultdict(list)
    try:
        coordinates = target_dict[type_of_pdf]
        for pg_num, pg in input_dict.items():
            page = doc[pg_num - 1]
            page_one = doc[0]
            for i, wlist in enumerate(pg):
                for k, target in coordinates.items():
                    if target and isinstance(target, tuple):
                        tuple_list = page_one.get_text("blocks", clip=target)
                        append_word_to_dict(tuple_list, field_dict[k], 0)
                    elif target[0] and isinstance(target[0], list) and any(target[0][0] in s for s in wlist[0]):
                        target_coords = target[0][1]
                        input_coords = input_dict[pg_num][i][1]
                        coords = tuple(x + y for x, y in zip(input_coords, target_coords))
                        word_list = page.get_text("blocks", clip=coords)
                        append_word_to_dict(word_list, field_dict[k], target[0][0][2])
                    elif target[0] and isinstance(target[0], str) and any(target[0] in s for s in wlist[0]):
                        word = input_dict[pg_num][i + target[1]][0][target[2]]
                        if target[0][3]:
                            field_dict[k].append(word)
                        if word and word not in field_dict[k]:
                            field_dict[k].append(word)
    except KeyError:
        return "Insurer Key does not exist"
    return field_dict

# 5 Clean dictionary:


def write_to_new_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["risk_type"].title()}.docx"
    doc.save(unique_file_name(output_path))


# 6 append to Pandas Dataframe:
def create_pandas_df(data_dict):
    df = pd.DataFrame([data_dict])
    filename = {"RENEWAL LETTER_NEW": "Renewal Letter - Copy.docx"}
    for rows in df.to_dict(orient="records"):
        if rows["risk_type"] == "Homeowners":
            write_to_new_docx(filename["RENEWAL LETTER_NEW"], rows)
        elif rows["risk_type"] == "Condominium":
            write_to_new_docx(filename["RENEWAL LETTER_NEW"], rows)
    return df


# Used for filling docx
def write_to_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    doc.save(unique_file_name(output_path))


# Used for fillable pdf's
def write_to_pdf(pdf, dictionary, rows):
    pdf_path = (base_dir / "templates" / pdf)
    output_path = base_dir / "output" / f"{rows["named_insured"]} {rows["risk_type"].title()}.pdf"
    output_path.parent.mkdir(exist_ok=True)
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    with open(unique_file_name(output_path), "wb") as output_stream:
        writer.write(output_stream)
