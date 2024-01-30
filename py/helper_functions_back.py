import os
from pathlib import Path



class CreateFile:

    base_dir = Path(__file__).parent.parent
    emoji = "\U0001F911\U0001F922\U0001F923\U0001F975\U0001F638"

    def __init__(self, filename, data):
        self.filename = filename
        self.data = data

    def output_dir(self, filetype):
        self.filetype = filetype
        output_dir = self.base_dir / "output" / Path(self.filename).stem
        output_dir.mkdir(exist_ok=True)
        dir_path = output_dir / f"{self.data}_{Path(self.filename).stem}.{self.filetype}"
        return dir_path

    def form_fields(self):
        self.output_dir()
        with open(dir_path, 'w', encoding="utf-8") as file:
            for page, value in enumerate(self.data):
            # for page, value in self.field_dict.items():
                file.write(f"\n<========= Page: {page} =========>\n")
                print(f"\n{self.emoji}   Page: {page} {self.emoji}\n")
                for text, box in enumerate(value):
                    file.write(f"#{text} : {box}\n")
                    print(f"#{text} :{box}")

class PageCoords(CreateTxtFile):

    def __init__(self, file_name, field_dict):
        super().__init__(file_name)
        self.field_dict = field_dict






def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path

def newline_to_list(nested_dict):
    for key, nested_list in nested_dict.items():
        if nested_list and isinstance(nested_list[0], list):
            for inner_list in nested_list:
                if inner_list:
                    inner_list[0] = inner_list[0].split('\n')
    return nested_dict

