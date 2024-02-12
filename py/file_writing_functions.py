import re
import pandas as pd
from collections import namedtuple
from datetime import datetime
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


# 3rd search to find the dictionary from the relevant pages
def search_for_input_dict(doc, pg_list):
    field_dict = {}
    for page_num in pg_list:
        page = doc[page_num - 1]
        wlist = page.get_text("blocks")
        text_boxes = [inner_list[4].split("\n") for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_num] = [[elem1, elem2] for elem1, elem2 in
                                zip(text_boxes, text_coords)]
    return field_dict


# 4th find the keys for each matching field and increment outer and inner index of nested list
# if the coordinates is a list it will open the doc to find the coords to increment to the target value
# if the coordinates is a string, it will increment by index relative of input coords to target coords
# The first way is more accurate, but need to copy and paste input and target coordinates
# the second method using index won't always work if the input index and target index is too far apart

def append_word_to_dict(wlist, dict):
    for words in wlist:
        word = words[4].strip().split("\n")
        if word not in dict:
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
                        append_word_to_dict(tuple_list, field_dict[k])
                    elif target[0] and isinstance(target[0], list) and any(target[0][0] in s for s in wlist[0]):
                        target_coords = target[0][1]
                        input_coords = input_dict[pg_num][i][1]
                        coords = tuple(x + y for x, y in zip(input_coords, target_coords))
                        word_list = page.get_text("blocks", clip=coords)
                        append_word_to_dict(word_list, field_dict[k])
                    elif target[0] and isinstance(target[0], str) and any(target[0] in s for s in wlist[0]):
                        my_list = [input_dict[pg_num][i + target[1]][0][target[2]]]
                        word = [x for x in my_list if x]
                        if word and word not in field_dict[k]:
                            field_dict[k].append(word)
    except KeyError:
        return "Insurer Key does not exist"
    return field_dict


# 5 Clean dictionary:

postal_code_regex = r"[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ -]?\d[ABCEGHJ-NPRSTV-Z]\d$"


def find_index_of_regex(strings_list, regex_pattern):
    return [index for index, string in enumerate(strings_list) if re.search(regex_pattern, string)]


def find_index_of_substrings(strings_list, substring):
    return [index for index, string in enumerate(strings_list) if substring in string]


def split_string_with_regex(pattern, strings):
    return re.split(pattern, strings)


def title_case(strings_list, length):
    if isinstance(strings_list, list):
        word = [string.title() if len(string) > length else string for string in strings_list]
        return word[0]
    if isinstance(strings_list, str):
        words = strings_list.split()
        capitalized_words = [word.capitalize() if len(word) > length else word for word in words]
        return ' '.join(capitalized_words)

def w2n(word):
    word_to_number = {
        "one": 1,
        "two": 2,
        "three": 3,
        # Add more mappings as needed
    }
    return word_to_number.get(word.strip().lower(), None)

def w2f(word):
    field_dict = {
        "COMPREHENSIVE FORM": "Comprehensive",
        # Add more mappings as needed
    }
    for key in field_dict.keys():
        if word.lower() in key.lower():
            return field_dict[word]
    else:
        return word

def w2r(word):
    field_dict = {
        "HOMEOWNERS": "Homeowners",
        "CONDOMINIUM": "Condominium",
        # Add more mappings as needed
    }
    for key in field_dict.keys():
        if word.lower() in key.lower():
            return field_dict[key]
    else:
        return word

def convert_to_int(amount_str):
    clean_amount_str = amount_str.replace("$", "").replace(",", "")
    return int(clean_amount_str)

def sum_dollar_amounts(amounts):
    # Convert each amount string to an integer and sum them up
    total = sum(convert_to_int(amount_str) for amount_str in amounts if amount_str)
    return total


def format_dict_items(dict_items, type_of_pdf):
    field_dict = {}
    if type_of_pdf == "Aviva":

        # Address block formatting:
        second_name_exists = find_index_of_substrings(dict_items["name_and_address"][0], "&")
        pc_index = find_index_of_regex(dict_items["name_and_address"][0], postal_code_regex)
        rpc_index = find_index_of_regex(dict_items["risk_address"][0][0], postal_code_regex)
        field_dict["named_insured"] = dict_items["name_and_address"][0][0].title()
        if second_name_exists:
            field_dict["addtional_insured"] = re.sub(r"&", "", dict_items["name_and_address"][0][1].title())
            field_dict["address_line_one"] = title_case(dict_items["name_and_address"][0][2:pc_index[0]], 1)
        else:
            field_dict["address_line_one"] = title_case(dict_items["name_and_address"][0][1:pc_index[0]], 1)
        city_province_postal = dict_items["name_and_address"][0][pc_index[0]]
        field_dict["address_line_two"] = title_case(re.sub(postal_code_regex, "", city_province_postal), 2)
        field_dict["address_line_three"] = re.search(postal_code_regex, city_province_postal).group()
        # Policy number formatting:
        field_dict["policy_number"] = re.sub(r"POLICY NUMBER: ", "", dict_items["policy_number"][0][0])
        # Effective date formatting:
        time_with_date = re.sub(r" 12:01 a.m.", "", dict_items["effective_date"][0][0])
        field_dict["effective_date"] = datetime.strptime(time_with_date, "%B %d, %Y").strftime("%B %d, %Y")
        # Risk address formatting:
        if rpc_index:
            for risk_address in dict_items["risk_address"][0]:
                field_dict["risk_address"] = " ".join(risk_address)
            if dict_items["location_2"]:
                for location_2 in dict_items["location_2"][0]:
                    field_dict["location_2"] = " ".join(location_2)
        # Form type formatting:
        field_dict["risk_type"] = w2r([i.split(" - ")[0] for i in dict_items["form_type"][0]][0])
        field_dict["form_type"] = w2f([i.split(" - ")[1] for i in dict_items["form_type"][0]][0])
        # Number of families formatting:
        if dict_items["number_families"]:
            number_of_families = dict_items["number_families"][0][0].split(" ,")
            afamily_index = find_index_of_regex(number_of_families, r"Family")
            field_dict["number_families"] = w2n(re.sub(r" Family", "", number_of_families[afamily_index[0]]))
        # Deductible formatting:
        field_dict["policy_deductible"] = re.search(r'\$([\d,]+)', dict_items["policy_deductible"][0][0]).group()
        # Condo deductible formatting:
        if dict_items["condo_deductible_coverage"]:
          field_dict["condo_deductible_coverage"] = re.search(r'\$([\d,]+)', dict_items["condo_deductible_coverage"][0][0]).group()
        # Premium amount formatting:
        field_dict["premium_amount"] = f"${sum_dollar_amounts(dict_items["premium_amount"][0][0])}"

    return field_dict
    # return field_dict

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
