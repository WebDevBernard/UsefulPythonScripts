import re
import pandas as pd
from collections import defaultdict
from coordinates import base_dir
from datetime import datetime
from coordinates import (postal_code_regex, dollar_regex, dict_of_keywords, date_regex, and_regex, address_regex)
from formatting_functions import (ff, remove_non_match, return_match_only, title_case, match_keyword,
                                  sum_dollar_amounts, flatten, find_index, find_nested_match, clean_dollar_amounts,
                                  unique_file_name)


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
                    for j in range(i + coords, i + 1):
                        pg_list.append(j + 1)
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
def append_word_to_dict(wlist, field_dict, ignore_duplicates):
    for words in wlist:
        word = words[4].strip().split("\n")
        if ignore_duplicates:
            field_dict.append(word)
        if word and word not in field_dict:
            field_dict.append(word)


def search_for_matches(doc, input_dict, type_of_pdf, target_dict):
    field_dict = defaultdict(lambda: defaultdict(list))
    try:
        coordinates = target_dict[type_of_pdf]
        for pg_num, pg in input_dict.items():
            page = doc[pg_num - 1]
            page_one = doc[0]
            for i, wlist in enumerate(pg):
                for k, target in coordinates.items():
                    if target and isinstance(target, tuple):
                        tuple_list = page_one.get_text("blocks", clip=target)
                        append_word_to_dict(tuple_list, field_dict[type_of_pdf][k], False)
                    try:
                        if target[0] and isinstance(target[0], list) and any(target[0][0] in s for s in wlist[0]):
                            target_coords = target[0][1]
                            input_coords = input_dict[pg_num][i][1]
                            coords = tuple(x + y for x, y in zip(input_coords, target_coords))
                            word_list = page.get_text("blocks", clip=coords)
                            append_word_to_dict(word_list, field_dict[type_of_pdf][k], target[0][0][2])
                        elif target[0] and isinstance(target[0], str) and any(target[0] in s for s in wlist[0]):
                            word = input_dict[pg_num][i + target[1]][0][target[2]]
                            if target[0][3]:
                                field_dict[type_of_pdf][k].append(word)
                            if word and word not in field_dict[type_of_pdf][k]:
                                field_dict[type_of_pdf][k].append(word)
                        elif target[0] and isinstance(target[0], re.Pattern):
                            word = input_dict[pg_num][i + target[1]][0][target[2]]
                            if return_match_only(target[0], word) and word not in field_dict[type_of_pdf][k]:
                                field_dict[type_of_pdf][k].append(word)
                    except IndexError:
                        continue
    except KeyError:
        return
    return field_dict


# 5 Clean dictionary:
def format_named_insured(field_dict, dict_items, type_of_pdf):
    try:
        address_index = find_index(address_regex, dict_items["name_and_address"])
        if type_of_pdf == "Intact":
            if dict_items["name_and_address"] and isinstance(dict_items["name_and_address"], list):
                names = [i.split(' & ') for i in dict_items["name_and_address"][:address_index]]
                join_same_last_names = [" ".join(reversed(i.split(", "))).title() if ", " in i else
                                        (i + " " + names[0][0]).split(", ")[0].title() for i in names[0]]
                if len(join_same_last_names) > 1:
                    join_same_last_names[-1] = "and " + join_same_last_names[-1]
                field_dict["named_insured"] = ", ".join(join_same_last_names)
        if type_of_pdf == "Aviva" or type_of_pdf == "Family" or type_of_pdf == "Wawanesa":
            if dict_items["name_and_address"] and isinstance(dict_items["name_and_address"], list):
                names = remove_non_match(and_regex, dict_items["name_and_address"][:address_index])
                names = [i.strip().title() for i in names]
                if len(names) > 1:
                    names[-1] = "and " + names[-1]
                field_dict["named_insured"] = ", ".join(names)
    except TypeError:
        return
    return field_dict


def format_insurer_name(field_dict, type_of_pdf):
    try:
        field_dict["insurer"] = type_of_pdf
    except TypeError:
        return
    return field_dict


def format_mailing_address(field_dict, dict_items, type_of_pdf):
    try:
        if type_of_pdf == "Aviva" or type_of_pdf == "Family" or type_of_pdf == "Intact" or type_of_pdf == "Wawanesa":
            if dict_items["name_and_address"] and isinstance(dict_items["name_and_address"], list):
                pc_index = find_index(postal_code_regex, dict_items["name_and_address"])
                address_index = find_index(address_regex, dict_items["name_and_address"])
                city_province_p_code = " ".join(dict_items["name_and_address"][address_index + 1:pc_index + 1])
                field_dict["address_line_one"] = title_case(dict_items["name_and_address"][address_index:pc_index], 1)
                field_dict["address_line_two"] = title_case(remove_non_match(postal_code_regex, city_province_p_code),
                                                            2)
                field_dict["address_line_three"] = title_case(
                    return_match_only(postal_code_regex, city_province_p_code), 3)
    except TypeError:
        return
    return field_dict


def format_risk_address(field_dict, dict_items, type_of_pdf):
    try:
        if type_of_pdf == "Intact":
            all_addresses = find_nested_match(postal_code_regex, flatten(dict_items["risk_address"]))
            if isinstance(dict_items["risk_address"], list):
                for index, address in enumerate(all_addresses):
                    field_dict[f"risk_address_{index + 1}"] = title_case(
                        remove_non_match(postal_code_regex, all_addresses[index]), 3).rstrip(", ")
        if type_of_pdf == "Family" or type_of_pdf == "Intact" or type_of_pdf == "Wawanesa":
            all_addresses = find_nested_match(postal_code_regex, dict_items["risk_address"])
            if isinstance(dict_items["risk_address"], list):
                for index, address in enumerate(all_addresses):
                    field_dict[f"risk_address_{index + 1}"] = title_case(
                        remove_non_match(postal_code_regex, all_addresses[index]), 3).rstrip(", ")
            else:
                field_dict["risk_address_1"] = title_case(
                    remove_non_match(postal_code_regex, dict_items["risk_address"]), 3).rstrip(", ")
        if type_of_pdf == "Aviva":
            if isinstance(dict_items["risk_address"][0], list):
                field_dict["risk_address_1"] = title_case(
                    remove_non_match(postal_code_regex, " ".join(dict_items["risk_address"][0])), 3).rstrip(", ")
            else:
                field_dict["risk_address_1"] = title_case(
                    remove_non_match(postal_code_regex, dict_items["risk_address"][0]), 3).rstrip(", ")
            if isinstance(dict_items["risk_address"][1], list):
                field_dict["risk_address_2"] = title_case(
                    remove_non_match(postal_code_regex, " ".join(dict_items["risk_address"][1])), 3).rstrip(", ")
    except AttributeError:
        return
    return field_dict


def format_form_type(field_dict, dict_items, type_of_pdf):
    try:
        if type_of_pdf == "Family":
            if "included".casefold() in dict_items["form_type"].casefold():
                field_dict["form_type_1"] = "Comprehensive Form"
            else:
                field_dict["form_type_1"] = "Specified Perils"
        if type_of_pdf == "Intact":
            if isinstance(dict_items["risk_address"], list):
                for index, form_type in enumerate(
                        find_nested_match(re.compile(r'\((.*?)\)'), flatten(dict_items["risk_address"]))):
                    if "comprehensive".casefold() in form_type.casefold():
                        field_dict[f"form_type_{index + 1}"] = "Comprehensive Form"
                    if "broad".casefold() in form_type.casefold():
                        field_dict[f"form_type_{index + 1}"] = "Broad Form"
                    if "basic".casefold() in form_type.casefold():
                        field_dict[f"form_type_{index + 1}"] = "Basic Form"
                    if "fire & extended".casefold() in form_type.casefold():
                        field_dict[f"form_type_{index + 1}"] = "Basic Form"
        if isinstance(dict_items["form_type"], str):
            if "comprehensive".casefold() in dict_items["form_type"].casefold():
                field_dict["form_type_1"] = "Comprehensive Form"
            if "broad".casefold() in dict_items["form_type"].casefold():
                field_dict["form_type_1"] = "Broad Form"
            if "basic".casefold() in dict_items["form_type"].casefold():
                field_dict["form_type_1"] = "Basic Form"
            if "fire & extended".casefold() in dict_items["form_type"].casefold():
                field_dict["form_type_1"] = "Basic Form"
        if isinstance(dict_items["form_type"], list):
            for index, form_type in enumerate(dict_items["form_type"]):
                if "comprehensive".casefold() in form_type.casefold():
                    field_dict[f"form_type_{index + 1}"] = "Comprehensive Form"
                if "broad".casefold() in form_type.casefold():
                    field_dict[f"form_type_{index + 1}"] = "Broad Form"
                if "basic".casefold() in form_type.casefold():
                    field_dict[f"form_type_{index + 1}"] = "Basic Form"
                if "fire & extended".casefold() in form_type.casefold():
                    field_dict[f"form_type_{index + 1}"] = "Basic Form"
    except TypeError:
         return
    return field_dict


def format_risk_type(field_dict, dict_items, type_of_pdf):
    try:
        if type_of_pdf == "Family":
            if "home".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "home"
            if "condo".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "condo"
        if isinstance(dict_items["risk_type"], str):
            if "seasonal".casefold() in dict_items["risk_type"].casefold():
                field_dict["seasonal"] = True
            if "home".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "home"
            if type_of_pdf == "Wawanesa":
                if "Condominium" in dict_items["risk_type"]:
                    field_dict["risk_type"] = "condo"
            elif type_of_pdf == "Aviva":
                if "condominium".casefold() in dict_items["risk_type"].casefold():
                    field_dict["risk_type"] = "condo"
            if "rented dwelling".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "rented_dwelling"
            if "revenue".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "rented_dwelling"
            if "rental".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "rented_condo"
            if "tenant".casefold() in dict_items["risk_type"].casefold():
                field_dict["risk_type"] = "tenant"
        if type_of_pdf == "Intact":
            list_of_risk_types = []
            if isinstance(dict_items["risk_address"], list):
                for index, risk_type in enumerate(
                        find_nested_match(re.compile(r'\((.*?)\)'), flatten(dict_items["risk_address"]))):
                    if "seasonal".casefold() in risk_type.casefold():
                        field_dict["seasonal"] = True
                    if "home".casefold() in risk_type.casefold():
                        list_of_risk_types.append("home")
                    if "condominium ".casefold() in risk_type.casefold():
                        list_of_risk_types.append("condo")
                    if "rented dwelling".casefold() in risk_type.casefold():
                        list_of_risk_types.append("rented_dwelling")
                    if "revenue".casefold() in risk_type.casefold():
                        list_of_risk_types.append("rented_dwelling")
                    if "rented condominium".casefold() in risk_type.casefold():
                        list_of_risk_types.append("rented_condo")
                    if any("home".casefold() in s.casefold() for s in list_of_risk_types):
                        field_dict["risk_type"] = "Home"
                    else:
                        try:
                            field_dict["risk_type"] = list_of_risk_types[0]
                        except IndexError:
                            return
        if isinstance(dict_items["risk_type"], list):
            list_of_risk_types = []
            for index, risk_type in enumerate(dict_items["risk_type"]):
                if "seasonal".casefold() in risk_type.casefold():
                    field_dict["seasonal"] = True
                if "home".casefold() in risk_type.casefold():
                    list_of_risk_types.append("home")
                if type_of_pdf == "Wawanesa":
                    if "Condominium" in risk_type:
                        field_dict["risk_type"] = "condo"
                elif type_of_pdf == "Aviva":
                    if "condominium".casefold() in risk_type.casefold():
                        field_dict["risk_type"] = "condo"
                if "rented dwelling".casefold() in risk_type.casefold():
                    list_of_risk_types.append("rented_dwelling")
                if "revenue".casefold() in risk_type.casefold():
                    list_of_risk_types.append("rented_dwelling")
                if "rental condominium".casefold() in risk_type.casefold():
                    list_of_risk_types.append("rented_condo")
                if any("home".casefold() in s.casefold() for s in list_of_risk_types):
                    field_dict["risk_type"] = "home"
                else:
                    try:
                        field_dict["risk_type"] = list_of_risk_types[0]
                    except IndexError:
                        return
    except TypeError:
         return
    return field_dict


def format_number_families(field_dict, dict_items, type_of_pdf):
    try:
        if type_of_pdf == "Family" and dict_items["number_of_families"]:
            for index, families in enumerate(dict_items["number_of_families"]):
                match = re.search(r"\b(\d+)\b", families)
                if match:
                    number = str(int(match.group(1)) + 1)
                    field_dict[f"number_of_families"] = match_keyword(dict_of_keywords, number)
        elif type_of_pdf == "Family":
            field_dict[f"number_of_families"] = match_keyword(dict_of_keywords, "1")
        if type_of_pdf == "Intact" or type_of_pdf == "Wawanesa":
            if isinstance(dict_items["number_of_families"], str) or isinstance(dict_items["number_of_units"], str):
                field_dict["number_of_families"] = match_keyword(dict_of_keywords, dict_items["number_of_families"])
                field_dict["number_of_families"] = match_keyword(dict_of_keywords, dict_items["number_of_units"])
        if type_of_pdf == "Aviva" and isinstance(dict_items["number_of_families"], str):
            field_dict["number_of_families"] = match_keyword(dict_of_keywords, remove_non_match(re.compile(
                r" Family"), dict_items["number_of_families"].split(" , ")[0]))
        elif type_of_pdf == "Aviva" and isinstance(dict_items["number_of_families"], list):
            for index, family_number in enumerate(dict_items["number_of_families"]):
                field_dict["number_of_families"] = match_keyword(dict_of_keywords, remove_non_match(re.compile(
                    r" Family"), family_number.split(" , ")[0]))
    except TypeError:
         return
    return field_dict


def format_policy_number(field_dict, dict_items):
    try:
        if isinstance(dict_items["policy_number"], list):
            for index, policy_number in enumerate(dict_items["number_of_families"]):
                field_dict["policy_number"] = policy_number
        field_dict["policy_number"] = dict_items["policy_number"]
    except TypeError:
         return
    return field_dict


def format_effective_date(field_dict, dict_items):
    try:
        time_with_date = return_match_only(date_regex, dict_items["effective_date"])
        field_dict["effective_date"] = time_with_date
    except TypeError:
        return
    return field_dict


def format_condo_deductible(field_dict, dict_items, type_of_pdf):
    try:
        if isinstance(dict_items["condo_deductible"], str):
            field_dict["condo_deductible_1"] = return_match_only(dollar_regex,
                                                                 dict_items["condo_deductible"])
        else:
            for index, condo_deductible in enumerate(dict_items["condo_deductible_coverage"]):
                field_dict[f"condo_deductible_{index + 1}"] = return_match_only(dollar_regex, condo_deductible)
        if type_of_pdf == "Family" and dict_items["condo_deductible"]:
            field_dict["condo_deductible_1"] = return_match_only(dollar_regex, dict_items["condo_deductible"][0])
    except TypeError:
         return
    return field_dict


def format_condo_earthquake_deductible(field_dict, dict_items, type_of_pdf):
    try:
        if isinstance(dict_items["condo_earthquake_deductible"], str):
            field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex,
                                                                            dict_items["condo_earthquake_deductible"])
        else:
            for index, condo_deductible in enumerate(dict_items["condo_earthquake_deductible"]):
                field_dict[f"condo_earthquake_deductible_{index + 1}"] = return_match_only(dollar_regex,
                                                                                           condo_deductible)
        if type_of_pdf == "Family" and dict_items["condo_deductible"]:
            field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex,
                                                                            dict_items["condo_deductible"][1])
        if type_of_pdf == "Intact" and dict_items["condo_earthquake_deductible"]:
            field_dict["condo_earthquake_deductible_1"] = "$25,000"
    except TypeError:
         return
    return field_dict


def format_premium_amount(field_dict, dict_items):
    try:
        if isinstance(dict_items["premium_amount"], list):
            field_dict["premium_amount"] = '${:,.2f}'.format(sum_dollar_amounts(dict_items["premium_amount"]))
        else:
            field_dict["premium_amount"] = '${:,.2f}'.format(clean_dollar_amounts(dict_items["premium_amount"]))
    except TypeError:
         return
    return field_dict


def format_additional_coverage(field_dict, dict_items, type_of_pdf):
    try:
        if (type_of_pdf == "Family" and dict_items["earthquake_coverage"] and
                return_match_only(dollar_regex, dict_items["earthquake_coverage"])):
            field_dict["earthquake_coverage"] = True
        if type_of_pdf == "Aviva" or type_of_pdf == "Intact" or type_of_pdf == "Wawanesa":
            if dict_items["earthquake_coverage"]:
                field_dict["earthquake_coverage"] = True
        if type_of_pdf == "Intact":
            if dict_items["ground_water"]:
                field_dict["ground_water"] = True
            field_dict["condo_deductible_1"] = "$100,000"
        if dict_items["overland_water"]:
            field_dict["overland_water"] = True
        if dict_items["service_line"]:
            field_dict["service_line"] = True
        if type_of_pdf == "Wawanesa":
            if dict_items["tenant_vandalism"]:
                field_dict["tenant_vandalism"] = True
    except TypeError:
        return


def format_policy(dict_items, type_of_pdf):
    field_dict = {}
    if type_of_pdf:
        flattened_dict = ff(dict_items[type_of_pdf])
        format_named_insured(field_dict, flattened_dict, type_of_pdf)
        format_insurer_name(field_dict, type_of_pdf)
        format_mailing_address(field_dict, flattened_dict, type_of_pdf)
        format_policy_number(field_dict, flattened_dict)
        format_effective_date(field_dict, flattened_dict)
        format_risk_address(field_dict, flattened_dict, type_of_pdf)
        format_form_type(field_dict, flattened_dict, type_of_pdf)
        format_risk_type(field_dict, flattened_dict, type_of_pdf)
        format_number_families(field_dict, flattened_dict, type_of_pdf)
        format_condo_deductible(field_dict, flattened_dict, type_of_pdf)
        format_condo_earthquake_deductible(field_dict, flattened_dict, type_of_pdf)
        format_premium_amount(field_dict, flattened_dict)
        format_additional_coverage(field_dict, flattened_dict, type_of_pdf)
    return field_dict

# 6 append to Pandas Dataframe:
def create_pandas_df(data_dict):
    df = pd.DataFrame([data_dict])
    df["today"] = datetime.today().strftime("%B %d, %Y")
    df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
    expiry_date = pd.to_datetime(df["effective_date"]) + pd.offsets.DateOffset(years=1)
    df["expiry_date"] = expiry_date.dt.strftime("%B %d, %Y")
    return df
