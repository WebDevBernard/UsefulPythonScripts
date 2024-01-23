from pathlib import Path
import pdfplumber
def remove_empty_strings(dictionary):
    # Iterate through each key-value pair in the dictionary
    for key, value in dictionary.items():
        # Check if the value is a list of tuples
        if isinstance(value, list):
            # Iterate through each tuple in the list
            updated_list = []
            for tpl in value:
                # Check if the tuple contains a list and another tuple
                if isinstance(tpl, tuple) and len(tpl) == 2 and isinstance(tpl[0], list) and isinstance(tpl[1], tuple):
                    # Remove empty strings from list 1
                    filtered_list1 = [item for item in tpl[0] if item not in ['', None]]
                    # Check if list 1 is not empty and update the tuple
                    if filtered_list1:
                        updated_list.append((filtered_list1, tpl[1]))
            # Update the dictionary with the modified list
            dictionary[key] = updated_list

    return dictionary
def extract_tables_from_pdf(pdf_path, strategy):
    table_settings = strategy
    result_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            tables = []
            list_of_tables = []     
            page = pdf.pages[page_number]
            im = page.to_image(resolution=400)
            im.debug_tablefinder(table_settings)
            if img_capture:
                im.show()
            find_table_object = page.find_tables(table_settings=table_settings)
            tables_on_page = page.extract_tables(table_settings=table_settings)
            for table in (find_table_object or []):
                list_of_tables.extend(table.cells)
            for table in (tables_on_page or []):
                tables.extend(table)
            result_dict[page_number + 1] = list(zip(tables, list_of_tables))
    return remove_empty_strings(result_dict)

base_dir = Path(__file__).parent.parent
pdf_directory = Path(base_dir / "input")
pdf_files = pdf_directory.glob("*.pdf")
pdf_file_paths = []
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

strategy = {
            # "vertical_strategy": "lines",
            # "horizontal_strategy": "text"
            }
img_capture = {
            # "True": "False"
            }

def filename_in_brackets():
    if bool(strategy):
        return "(horizontal_strategy_text)"
    else:
        return "(horizontal_strategy_lines)" 

# Loop through each PDF file and append the full path to the list
for pdf_file in pdf_files:
    file_path = str(pdf_file)
    pdf_file_paths.append(file_path)

for pdf_path in pdf_file_paths:
    print(f"\n<========= PDF_FILENAME: {pdf_path} =========>\n")
    pages = extract_tables_from_pdf(pdf_path, strategy)
    with open( output_dir / f"{Path(pdf_path).stem} Table Coordinates {filename_in_brackets()}.txt", 'w') as file:
        for page, value in pages.items():
            file.write(f"\n<========= Page: {page} =========>\n")
            print(f"\n<========= Page: {page} =========>\n")
            for text, box in enumerate(value):
                file.write(f"#{text} : {box}\n")
                print(f"#{text} : {box}")
