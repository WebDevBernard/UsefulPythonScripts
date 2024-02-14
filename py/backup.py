




def is_nested_list(lists):
    try:
        next(x for x in lists if isinstance(x, list))
    except StopIteration:
        return False
    return True


def find_index(strings_list, substring):
    if is_nested_list(strings_list):
        return [index for index, string in enumerate(strings_list[0]) if re.search(substring, string)][0]
    else:
        return [index for index, string in enumerate(strings_list) if re.search(substring, string)][0]


def title_case(strings_list, str_length):
    if strings_list and isinstance(strings_list, list):
        word = [string.strip().title() if len(string) > str_length else string for string in strings_list]
        return word[0]
    if strings_list and isinstance(strings_list, str):
        words = strings_list.split()
        capitalized_words = [word.strip().capitalize() if len(word) > str_length else word for word in words]
        return ' '.join(capitalized_words)


def assign_keyword(dict_of_keywords, keyword):
    return dict_of_keywords.get(keyword, None)


def sum_dollar_amounts(amounts):
    clean_amount_str = amounts.replace("$", "").replace(",", "")
    total = sum(int(clean_amount_str) for amount_str in amounts if amount_str)
    return total


def format_name_address(field_dict, dict_items):
    pc_index = find_index(dict_items["name_and_address"], postal_code_regex)
    second_name_exists = [index for index, string in enumerate(dict_items["name_and_address"]) if "&" in string][0]
    field_dict["named_insured"] = dict_items["name_and_address"][0].title()
    if second_name_exists:
        field_dict["addtional_insured"] = re.sub(r"&", "", dict_items["name_and_address"][1]).title()
        field_dict["address_line_one"] = title_case(dict_items["name_and_address"][2:pc_index], 1)
    else:
        field_dict["address_line_one"] = title_case(dict_items["name_and_address"][1:pc_index], 1)
    city_province_postal = dict_items["name_and_address"][pc_index]
    field_dict["address_line_two"] = title_case(re.sub(postal_code_regex, "", city_province_postal), 2)
    field_dict["address_line_three"] = re.search(postal_code_regex, city_province_postal).group()
    return field_dict





def format_effective_date(field_dict, dict_items):
    time_with_date = re.sub(r" 12:01 a.m.", "", dict_items["effective_date"])
    field_dict["effective_date"] = datetime.strptime(time_with_date, "%B %d, %Y").strftime("%B %d, %Y")
    return field_dict


def format_risk_address(field_dict, dict_items):
    if isinstance(dict_items["form_type"], list):
        risk_pc_index = find_index(dict_items["risk_address"], postal_code_regex)
        if risk_pc_index:
            for risk_address in dict_items["risk_address"]:
                field_dict["risk_address"] = " ".join(risk_address)
    # field_dict["location_2"] = dict_items["location_2"][risk_pc_index]
    return field_dict


# Location one and two can have different form types
def format_form_type(field_dict, dict_items, dict_of_keywords):
    if isinstance(dict_items["form_type"], list):
        field_dict["risk_type"] = assign_keyword(dict_of_keywords, dict_items["form_type"][0].split(" - ")[0])
        field_dict["form_type"] = assign_keyword(dict_of_keywords, dict_items["form_type"][0].split(" - ")[1])
        field_dict["risk_type_2"] = assign_keyword(dict_of_keywords, dict_items["form_type"][1].split(" - ")[0])
        field_dict["form_type_2"] = assign_keyword(dict_of_keywords, dict_items["form_type"][1].split(" - ")[1])
    else:
        field_dict["risk_type"] = assign_keyword(dict_of_keywords, dict_items["form_type"].split(" - ")[0])
        field_dict["form_type"] = assign_keyword(dict_of_keywords, dict_items["form_type"].split(" - ")[1])
    return field_dict

# this doesnt exists for condo so it won't work
def format_number_families(field_dict, dict_items, dict_of_keywords):
    if dict_items["number_families"]:
        if isinstance(dict_items["number_families"], list):
            number_of_families = dict_items["number_families"][0].split(" ,")
            number_of_families_2 = dict_items["number_families"][0].split(" ,")
            family_index = find_index(number_of_families, re.compile(r"Family"))
            field_dict["number_families"] = assign_keyword(dict_of_keywords,
                                                           re.sub(r" Family", "", number_of_families[family_index]))
            field_dict["number_families_2"] = assign_keyword(dict_of_keywords,
                                                           re.sub(r" Family", "", number_of_families[family_index]))
        else:
            number_of_families = dict_items["number_families"].split(" ,")
            family_index = find_index(number_of_families, re.compile(r"Family"))
            field_dict["number_families"] = assign_keyword(dict_of_keywords,
                                                           re.sub(r" Family", "", number_of_families[family_index]))
    return field_dict


def format_policy_deductible(field_dict, dict_items):
    field_dict["policy_deductible"] = re.search(r'\$([\d,]+)', dict_items["policy_deductible"]).group()
    return field_dict


def format_condo_deductible(field_dict, dict_items):
    if dict_items["condo_deductible_coverage"]:
        field_dict["condo_deductible_coverage"] = re.search(r'\$([\d,]+)',
                                                            dict_items["condo_deductible_coverage"]).group()
    return field_dict


def format_premium_amount(field_dict, dict_items):
    field_dict["premium_amount"] = '${:,.2f}'.format(sum_dollar_amounts(dict_items["premium_amount"][0]))
    return field_dict