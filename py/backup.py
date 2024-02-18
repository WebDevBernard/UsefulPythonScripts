import re
import pandas as pd
from collections import defaultdict
from debugging import base_dir, unique_file_name
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfWriter
from coordinates import (postal_code_regex, dollar_regex, dict_of_keywords, date_regex, and_regex, address_regex,
                         postal_code_regex_2, intact_date_regex)
from formatting_functions import (ff, remove_non_match, return_match_only, title_case, match_keyword, sum_dollar_amounts,
                                  find_index, find_nested_match, add_dollar_sign, find_nested_match_2)

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
            elif isinstance(kw, list):
                stop_word = page.search_for(kw[0])
                if stop_word:
                    pg_list.append(i + 1)
                    pg_list.append(i + 1 + coords)
                    break
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
def append_word_to_dict(wlist, dict, replace_duplicates):
    for words in wlist:
        word = words[4].strip().split("\n")
        if word and not replace_duplicates and word not in dict:
            dict.append(word)
        else:
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
                        append_word_to_dict(tuple_list, field_dict[k], False)
                    try:
                        if target[0] and isinstance(target[0], list) and any(target[0][0] in s for s in wlist[0]):
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
                        elif target[0] and isinstance(target[0], re.Pattern):
                            word = input_dict[pg_num][i + target[1]][0][target[2]]
                            if return_match_only(target[0], word) and word not in field_dict[k]:
                                field_dict[k].append(word)
                    except IndexError:
                        print("Target index does not exist")
    except KeyError:
        return "Insurer Key does not exist"
    return field_dict

# 5 Clean dictionary:
def format_intact(field_dict, dict_items):
    all_addresses = find_nested_match_2(r'\((.*?)\)', dict_items["intact_risk_address"])
    risk_address = find_nested_match_2(postal_code_regex, dict_items["intact_risk_address"])
    if dict_items["intact_name_and_address"] and isinstance(dict_items["name_and_address"], list):
        address_index = find_index(address_regex, dict_items["intact_name_and_address"])
        names = [i.split(' & ') for i in dict_items["intact_name_and_address"][:address_index]]
        names_2 = [" ".join(reversed(i.split(", "))).title() if ", " in i else
                   (i + " " + names[0][0]).split(", ")[0].title() for i in names[0]]
        field_dict["named_insured"] = ", ".join(names_2)
    if isinstance(dict_items["intact_risk_address"], list):
        for index, risk_type in enumerate(all_addresses):
            form_type = return_match_only(re.compile(r"\((.*?)\)"), all_addresses[index])
            risk_type = return_match_only(re.compile(r"\((.*?)\)"), all_addresses[index])
            try:
                risk_type_2 = match_keyword(dict_of_keywords, risk_type)[0]
                form_type_2 = match_keyword(dict_of_keywords, form_type)[1]
                if risk_type_2 == "Condo" or risk_type_2 == "Rental Condo":
                    field_dict["condo_deductible"] = "$100,000"
                field_dict[f"risk_type_{index + 1}"] = risk_type_2
                field_dict[f"form_type_{index + 1}"] = form_type_2
                field_dict[f"risk_address_{index + 1}"] = title_case(risk_address[index], 3)
                if dict_items["intact_condo_earthquake_deductible"]:
                    field_dict["condo_earthquake_deductible"] = "$25,000"
                if dict_items["intact_earthquake_coverage"]:
                    field_dict["earthquake_coverage"] = True
                if dict_items["intact_overland_water}"]:
                    field_dict["ground_water"] = True
                if dict_items["intact_ground_water"]:
                    field_dict["overland_water"] = True
            except TypeError:
                "Intact Key error"

    return field_dict


def format_named_insured(field_dict, dict_items):
    if dict_items["name_and_address"] and isinstance(dict_items["name_and_address"], list):
        address_index = find_index(address_regex, dict_items["name_and_address"])
        names = remove_non_match(and_regex, dict_items["name_and_address"][:address_index])
        names = [i.strip().title() for i in names]
        # if len(names) > 1:
        #     names[-1] = "and " + names[-1]
        field_dict["named_insured"] = ", ".join(names)


def format_mailing_address(field_dict, dict_items):
    if dict_items["name_and_address"] and isinstance(dict_items["name_and_address"], list):
        pc_index = find_index(postal_code_regex, dict_items["name_and_address"])
        address_index = find_index(address_regex, dict_items["name_and_address"])
        city_province_p_code = " ".join(dict_items["name_and_address"][address_index + 1:pc_index + 1])
        field_dict["address_line_one"] = title_case(dict_items["name_and_address"][address_index:pc_index], 1)
        field_dict["address_line_two"] = title_case(remove_non_match(postal_code_regex_2, city_province_p_code), 2)
        field_dict["address_line_three"] = title_case(return_match_only(postal_code_regex, city_province_p_code), 3)
    return field_dict


def format_risk_address(field_dict, dict_items):
    all_addresses = find_nested_match(postal_code_regex, dict_items["risk_address"])
    if isinstance(dict_items["risk_address"], str):
        field_dict["risk_address_1"] = dict_items["risk_address"]
    if isinstance(dict_items["risk_address"], list):
        for index, address in enumerate(all_addresses):
            field_dict[f"risk_address_{index + 1}"] = all_addresses[index]
    return field_dict


def format_form_type(field_dict, dict_items):
    if isinstance(dict_items["family_risk_type"], str):
        field_dict["risk_type_1"] = match_keyword(dict_of_keywords, dict_items["family_risk_type"])
    if isinstance(dict_items["wawa_risk_type"], list):
        for index, risk_type in enumerate(dict_items["wawa_risk_type"]):
            field_dict[f"risk_type_{index + 1}"] = match_keyword(dict_of_keywords, risk_type)
    if isinstance(dict_items["wawa_form_type"], str):
        field_dict[f"form_type_1"] = match_keyword(dict_of_keywords, dict_items["wawa_form_type"].split(" - ")[0])
    if isinstance(dict_items["wawa_form_type"], list):
        for index, form_type in enumerate(dict_items["wawa_form_type"]):
            field_dict[f"form_type_{index + 1}"] = match_keyword(dict_of_keywords, form_type.split(" - ")[0])
    if isinstance(dict_items["aviva_form_type"], str):
        field_dict["risk_type_1"] = match_keyword(dict_of_keywords, dict_items["aviva_form_type"].split(" - ")[0])
        field_dict["form_type_1"] = match_keyword(dict_of_keywords, dict_items["aviva_form_type"].split(" - ")[1])
    else:
        for index, risk_type in enumerate(dict_items["aviva_form_type"]):
            field_dict[f"risk_type_{index + 1}"] = match_keyword(dict_of_keywords, risk_type.split(" - ")[0])
            field_dict[f"form_type_{index + 1}"] = match_keyword(dict_of_keywords, risk_type.split(" - ")[1])
    return field_dict


def format_number_families(field_dict, dict_items):
    if dict_items["family_number_of_families"]:
        for index, families in enumerate(dict_items["family_number_of_families"]):
            match = re.search(r"\b(\d+)\b", families)
            if match:
                number = str(int(match.group(1)) + 1)
                field_dict[f"number_of_family_1"] = match_keyword(dict_of_keywords, number)
    else:
        field_dict[f"number_of_family_1"] = match_keyword(dict_of_keywords, "1")
    return field_dict


def format_policy_number(field_dict, dict_items):
    if isinstance(dict_items["policy_number"], str):
        field_dict["policy_number"] = dict_items["policy_number"]
    else:
        for index, policy_number in enumerate(dict_items["policy_number"]):
            field_dict["policy_number"] = policy_number
    return field_dict


def format_effective_date(field_dict, dict_items):
    if isinstance(dict_items["intact_effective_date"], str):
        time_with_date = return_match_only(intact_date_regex, dict_items["intact_effective_date"])
        field_dict["effective_date"] = time_with_date
    if isinstance(dict_items["intact_effective_date"], list):
        for index, time_date in enumerate(dict_items["intact_effective_date"]):
            time_with_date = return_match_only(intact_date_regex, time_date)
            field_dict["effective_date"] = time_with_date
    if isinstance(dict_items["effective_date"], str):
        time_with_date = return_match_only(date_regex, dict_items["effective_date"])
        field_dict["effective_date"] = time_with_date
    else:
        for index, time_date in enumerate(dict_items["effective_date"]):
            time_with_date = return_match_only(date_regex, time_date)
            field_dict["effective_date"] = time_with_date
    return field_dict


def format_condo_deductible(field_dict, dict_items):
    if isinstance(dict_items["condo_deductible"], str):
        field_dict["condo_deductible_1"] = return_match_only(dollar_regex,
                                                                      dict_items["condo_deductible"])
    else:
        for index, condo_deductible in enumerate(dict_items["condo_deductible_coverage"]):
            field_dict[f"condo_deductible_{index + 1}"] = return_match_only(dollar_regex, condo_deductible)
    if isinstance(dict_items["family_condo_deductible"], str):
        field_dict["condo_deductible_1"] = return_match_only(dollar_regex,
                                                                      dict_items["family_condo_deductible"])
    elif dict_items["family_condo_deductible"]:
        field_dict["condo_deductible_1"] = return_match_only(dollar_regex, dict_items["family_condo_deductible"][0])
    return field_dict


def format_condo_earthquake_deductible(field_dict, dict_items):
    if isinstance(dict_items["condo_earthquake_deductible"], str):
        field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex,
                                                                      dict_items["condo_earthquake_deductible"])
    else:
        for index, condo_deductible in enumerate(dict_items["condo_earthquake_deductible"]):
            field_dict[f"condo_earthquake_deductible_{index + 1}"] = return_match_only(dollar_regex, condo_deductible)
    if isinstance(dict_items["family_condo_deductible"], str):
        field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex,
                                                                      dict_items["family_condo_deductible"])
    elif dict_items["family_condo_deductible"]:
        field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex, dict_items["family_condo_deductible"][1])
    return field_dict

def format_premium_amount(field_dict, dict_items):
    if isinstance(dict_items["premium_amount"], list):
        field_dict["premium_amount"] = '${:,.0f}'.format(sum_dollar_amounts(dict_items["premium_amount"]))
    else:
        field_dict["premium_amount"] = add_dollar_sign(dict_items["premium_amount"])
    return field_dict

def format_overland_water(field_dict, dict_items):
    if dict_items["overland_water"]:
        field_dict["overland_water"] = True

def format_service_line(field_dict, dict_items):
    if dict_items["service_line"]:
        field_dict["service_line"] = True

def format_earthquake_coverage(field_dict, dict_items):
    if dict_items["earthquake_coverage"] and isinstance(dict_items["earthquake_coverage"], list):
        field_dict["earthquake_coverage"] = True
    if dict_items["family_earthquake_coverage"] and return_match_only(dollar_regex, dict_items["family_earthquake_coverage"]):
        field_dict["earthquake_coverage"] = True
    elif dict_items["family_earthquake_coverage"] and return_match_only(re.compile(r"Excluded", flags=re.IGNORECASE), dict_items["family_earthquake_coverage"]):
        field_dict["earthquake_coverage"] = False



def format_policy(dict_items):
    field_dict = {}
    flattened_dict = ff(dict_items)
    format_named_insured(field_dict, flattened_dict)
    format_mailing_address(field_dict, flattened_dict)
    format_policy_number(field_dict, flattened_dict)
    format_effective_date(field_dict, flattened_dict)
    format_risk_address(field_dict, flattened_dict)
    format_form_type(field_dict, flattened_dict)
    format_number_families(field_dict, flattened_dict)
    format_condo_deductible(field_dict, flattened_dict)
    format_condo_earthquake_deductible(field_dict, flattened_dict)
    format_premium_amount(field_dict, flattened_dict)
    format_overland_water(field_dict, dict_items)
    format_earthquake_coverage(field_dict, dict_items)
    format_service_line(field_dict, dict_items)
    format_intact(field_dict, dict_items)
    # print(flattened_dict)
    return field_dict


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
