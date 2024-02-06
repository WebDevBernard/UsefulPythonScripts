import os
from pathlib import Path

base_dir = Path(__file__).parent.parent

def write_text_coords(file_name, block_dict, table_dict, word_dict):
    output_dir = base_dir / "output" / Path(file_name).stem
    output_dir.mkdir(exist_ok=True)
    block_dir_path = output_dir / f"block_coordinates_{Path(file_name).stem}.txt"
    table_dir_path = output_dir / f"table_coordinates_{Path(file_name).stem}.txt"
    word_dir_path = output_dir / f"word_coordinates_{Path(file_name).stem}.txt"
    if block_dict:
        with open(block_dir_path, 'w', encoding="utf-8") as file:
            for page, value in block_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")
    if table_dict:
        with open(table_dir_path, 'w', encoding="utf-8") as file:
            for page, value in table_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")
    if word_dict:
        with open(word_dir_path, 'w', encoding="utf-8") as file:
            for page, value in word_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")


def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


def newline_to_dict(nested_dict):
    for key, nested_list in nested_dict.items():
        if nested_list and isinstance(nested_list, list):
            for inner_list in nested_list:
                if isinstance(inner_list, str):
                    inner_list[0] = inner_list[0].strip().split('\n')
                else:
                    inner_list[0] = inner_list[0].split('\n')
    return nested_dict
