import re
import os
from pathlib import Path

def flatten(xss):
    return [x for xs in xss for x in xs]

def ff(dictionary):
    for key, value in dictionary.items():
        if isinstance(value, list) and len(value) == 1:
            dictionary[key] = value[0]
        if isinstance(value, list) and len(value[0]) == 1:
            dictionary[key] = value[0][0]
    return dictionary


def find_index(regex, dict_item):
    if dict_item and isinstance(dict_item, list):
        for index, string in enumerate(dict_item):
            if re.search(regex, string):
                return index
        return -1

def find_nested_match(regex, nested_list):
    matched_lists = []
    for item in nested_list:
        if isinstance(item, str) and re.search(regex, item) and item not in matched_lists:
            matched_lists.append(item)
    return matched_lists

def return_match_only(regex, dict_item):
    try:
        if re.search(regex, dict_item) is not None:
            return re.search(regex, dict_item).group()
    except KeyError:
        "Mapping key does not exist"


def remove_non_match(regex, dict_item):
    if isinstance(dict_item, list):
        return eval(re.sub(regex, "", str(dict_item)))
    else:
        return re.sub(regex, "", dict_item)

def match_keyword(dict_of_keywords, keyword):
    return dict_of_keywords.get(keyword, None)

def custom_title_case(sentence):
    return ' '.join(
        word if word.isdigit() or word[-2:] in {"th", "rd"} else word.capitalize()
        for word in sentence.split()
    )

def title_case(strings_list, str_length):
    if strings_list and isinstance(strings_list, list):
        word = [string.strip().title() if len(string) > str_length else string for string in strings_list]
        return custom_title_case(word[0])
    if strings_list and isinstance(strings_list, str):
        words = strings_list.split()
        capitalized_words = [word.strip().capitalize() if len(word) > str_length else word for word in words]
        return ' '.join(capitalized_words)


def sum_dollar_amounts(amounts):
    clean_amount_str = [a.replace("$", "").replace(",", "") for a in amounts]
    total = sum(int(c) for c in clean_amount_str)
    return total

def clean_dollar_amounts(amounts):
    try:
        clean_amount_str = amounts.replace("$", "").replace(",", "").replace(" 00", "").replace(".00", "")
        total = int(clean_amount_str)
        return total
    except ValueError:
        "fatal error is fatal"

def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path