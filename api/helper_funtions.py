import os
from pathlib import Path

emoji = "\U0001F923\U0001F923\U0001F923\U0001F923\U0001F923"

def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path

def remove_newlines(nested_dict):
    for key, nested_list in nested_dict.items():
        if nested_list and isinstance(nested_list[0], list):
            for inner_list in nested_list:
                if inner_list:
                    inner_list[0] = inner_list[0].replace('\n', '')
    return nested_dict