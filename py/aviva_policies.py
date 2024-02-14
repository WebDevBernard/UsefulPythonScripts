import re
from datetime import datetime
from coordinates import postal_code_regex, dollar_regex, dict_of_keywords
from format_functions import (ff, remove_non_match, return_match_only, title_case, match_keyword, sum_dollar_amounts,
                              find_index, find_nested_match)

# 5 Clean dictionary:
def format_name_address(field_dict, dict_items):
    pc_index = find_index(postal_code_regex, dict_items["name_and_address"])
    second_name_exists = [index for index, string in enumerate(dict_items["name_and_address"]) if "&" in string]
    field_dict["named_insured"] = dict_items["name_and_address"][0].title()
    if second_name_exists:
        field_dict["additional_insured"] = remove_non_match(r"&", dict_items["name_and_address"][1]).title()
        field_dict["address_line_one"] = title_case(dict_items["name_and_address"][2:pc_index], 1)
    else:
        field_dict["address_line_one"] = title_case(dict_items["name_and_address"][1:pc_index], 1)
    city_province_postal = dict_items["name_and_address"][pc_index]
    field_dict["address_line_two"] = title_case(remove_non_match(postal_code_regex, city_province_postal), 2)
    field_dict["address_line_three"] = return_match_only(postal_code_regex, city_province_postal)
    return field_dict

def format_risk_address(field_dict, dict_items):
    all_addresses = find_nested_match(postal_code_regex, dict_items["risk_address"])
    for index, address in enumerate(all_addresses):
        field_dict[f"risk_address_{index + 1}"] = " ".join(all_addresses[index])
    return field_dict

# Location one and two can have different form types
def format_form_type(field_dict, dict_items):
    if isinstance(dict_items["aviva_form_type"], str):
        field_dict["risk_type_1"] = match_keyword(dict_of_keywords, dict_items["aviva_form_type"].split(" - ")[0])
        field_dict["form_type_1"] = match_keyword(dict_of_keywords, dict_items["aviva_form_type"].split(" - ")[1])
    else:
        for index, risk_type in enumerate(dict_items["aviva_form_type"]):
            field_dict[f"risk_type_{index + 1}"] = match_keyword(dict_of_keywords, risk_type.split(" - ")[0])
            field_dict[f"form_type_{index + 1}"] = match_keyword(dict_of_keywords, risk_type.split(" - ")[1])
    return field_dict


def format_number_families(field_dict, dict_items):
    if isinstance(dict_items["aviva_number_of_families"], str):
        field_dict["number_of_family_1"] = match_keyword(dict_of_keywords, remove_non_match(
            r" Family", dict_items["aviva_number_of_families"].split(", ")[0]).strip())
    else:
        for index, families in enumerate(dict_items["aviva_number_of_families"]):
            field_dict[f"number_of_family_{index + 1}"] = match_keyword(dict_of_keywords, remove_non_match(
                r" Family", families.split(", ")[0]).strip())
    return field_dict

def format_policy_number(field_dict, dict_items):
    field_dict["policy_number"] = dict_items["policy_number"]
    return field_dict

def format_effective_date(field_dict, dict_items):
    time_with_date = remove_non_match(r" 12:01 a.m.", dict_items["effective_date"])
    field_dict["effective_date"] = datetime.strptime(time_with_date, "%B %d, %Y").strftime("%B %d, %Y")
    return field_dict

def format_policy_deductible(field_dict, dict_items):
    item_exists = dict_items["policy_deductible"]
    if item_exists:
        field_dict["policy_deductible"] = return_match_only(dollar_regex, item_exists)
    return field_dict

def format_condo_deductible(field_dict, dict_items):
    if isinstance(dict_items["condo_deductible_coverage"], str):
        field_dict["condo_deductible_coverage_1"] = return_match_only(dollar_regex, dict_items["condo_deductible_coverage"])
    else:
        for index, condo_deductible in enumerate(dict_items["condo_deductible_coverage"]):
            field_dict[f"condo_deductible_coverage_{index + 1}"] = return_match_only(dollar_regex, condo_deductible)
    return field_dict

def format_premium_amount(field_dict, dict_items):
    item_exists = dict_items["premium_amount"]
    if item_exists:
        if isinstance(item_exists, list):
            field_dict["premium_amount"] = '${:,.0f}'.format(sum_dollar_amounts(item_exists))
        else:
            field_dict["premium_amount"] = item_exists
    return field_dict

def format_overland_water(field_dict, dict_items):
    if dict_items["overland_water"]:
        field_dict["overland_water"] = True

def format_earthquake_coverage(field_dict, dict_items):
    if dict_items["earthquake_coverage"]:
        field_dict["earthquake_coverage"] = True

def format_aviva_policy(dict_items):
    field_dict = {}
    flattened_dict = ff(dict_items)
    format_name_address(field_dict, flattened_dict)
    format_policy_number(field_dict, flattened_dict)
    format_effective_date(field_dict, flattened_dict)
    format_risk_address(field_dict, flattened_dict)
    format_form_type(field_dict, flattened_dict)
    format_number_families(field_dict, flattened_dict)
    format_policy_deductible(field_dict, flattened_dict)
    format_condo_deductible(field_dict, flattened_dict)
    format_premium_amount(field_dict, flattened_dict)
    format_overland_water(field_dict, dict_items)
    format_earthquake_coverage(field_dict, dict_items)
    # print(flattened_dict)
    return field_dict




