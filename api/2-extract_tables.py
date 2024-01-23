from pathlib import Path
import pdfplumber

def extract_tables_from_pdf(pdf_path):
    table_settings = {
        # "vertical_strategy": "lines",
        # "horizontal_strategy": "text",
    }
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            tables = []
            for page_number in range(len(pdf.pages)):
                page = pdf.pages[page_number]
                im = page.to_image(resolution=400)
                im.debug_tablefinder(table_settings)
                # im.show()
                tables_on_page = page.extract_tables(table_settings=table_settings)
                for table_number, table in enumerate(tables_on_page or []):
                    tables.extend(table)
            return [x for x in [list(filter(None, lst)) for lst in tables]if x]

base_dir = Path(__file__).parent.parent
pdf_directory = Path(base_dir / "input")
pdf_files = pdf_directory.glob("*.pdf")
pdf_file_paths = []
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

# Loop through each PDF file and append the full path to the list
for pdf_file in pdf_files:
    file_path = str(pdf_file)
    pdf_file_paths.append(file_path)

for pdf_path in pdf_file_paths:
    print(f"\n")
    print(f"<========= PDF_FILENAME: {pdf_path} =========>")
    print(f"\n")
    # gets table values
    tables = extract_tables_from_pdf(pdf_path)
    for row_number, row in enumerate(tables):
        print(f"Row {row_number + 1}: {row}")

    # bb_1 = (46.5600015, 143.63243797500002, 494.03997825, 153.8975676)
    # bb_2 = (46.5600015, 156.67647284999998, 494.03997825, 166.61757285)
    # bb_3 = (46.5600015, 168.67647284999998, 494.03997825, 178.61757285)
    # result = extract_text_within_bbox(pdf_path, bb_1)
    # result_2 = extract_text_within_bbox(pdf_path, bb_2)
    # result_3 = extract_text_within_bbox(pdf_path, bb_3)
    # print(result)
    # print(result_2)
    # print(result_3)