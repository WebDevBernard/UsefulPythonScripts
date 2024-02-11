import os
from pathlib import Path

base_dir = Path(__file__).parent.parent

def calculate_target_coords(input_coords, target_coords, direction, debug):
    if direction == "up":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_y_0 = y0_1 + y0_2 - y0_1
        debug_result_y_1 = y1_1 + y1_2 - y1_1
        result_y_0 = y0_2 - y0_1
        result_y_1 = y1_2 - y1_1
        if debug:
            return (x0_2, debug_result_y_0, x1_2, debug_result_y_1)
        else:
            return (0, result_y_0, 0, result_y_1)

    if direction == "down":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_y_0 = y0_1 + y0_2 - y0_1
        debug_result_y_1 = y1_1 + y1_2 - y1_1
        result_y_0 = y0_2 - y0_1
        result_y_1 = y1_2 - y1_1
        if debug:
            return (x0_2, debug_result_y_0, x1_2, debug_result_y_1)
        else:
            return (0, result_y_0, 0, result_y_1)
    if direction == "left":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_x_0 = x0_1 + x0_2 - x0_1
        debug_result_x_1 = x1_1 + x1_2 - x1_1
        result_x_0 = x0_1 - x0_2
        result_x_1 = x1_1 - x1_2
        if debug:
            return (debug_result_x_0, y0_1, debug_result_x_1, y1_1)
        else:
            return (result_x_0, 0, result_x_1, 0)
    if direction == "right":
        x0_1, y0_1, x1_1, y1_1 = input_coords
        x0_2, y0_2, x1_2, y1_2 = target_coords
        debug_result_x_0 = x0_1 + x0_2 - x0_1
        debug_result_x_1 = x1_1 + x1_2 - x1_1
        result_x_0 = x0_2 - x0_1
        result_x_1 = x1_2 - x1_1
        if debug:
            return (debug_result_x_0, y0_1, debug_result_x_1, y1_1)
        else:
            return (result_x_0, 0, result_x_1, 0)

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
