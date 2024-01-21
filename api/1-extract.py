from pathlib import Path
import pdfplumber

def find_keyword_in_pdfs(pdf_paths, keyword):
    matching_files = []
    for pdf_path in pdf_paths:
        pdf_path = Path(pdf_path)
        with pdfplumber.open(pdf_path) as pdf:
            for page_number in range(len(pdf.pages)):
                text = pdf.pages[page_number].extract_text()
                if keyword.lower() in text.lower():
                    matching_files.append(pdf_path)
                    break  # Break the loop if the keyword is found in any page
    return matching_files


def get_text_bounding_boxes(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text_boxes = []
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            left = page.crop((0, 0, 287, 771.412), relative=False, strict=True)
            right = page.crop((0.5 * float(page.width), 0.4 * float(page.height), page.width, 0.9 * float(page.height)))
            # left_text = left.extract_text_lines()
            right_text = right.extract_text_lines()
            text_lines = page.extract_text_lines()

            for line in text_lines:
            # for line in left_text:
                # Check if the line intersects with the specified region
                if line['text']:  # Ignore empty lines
                    bounding_box = {
                        'x0': ['x0'],
                        'y0': line['top'],
                        'x1': line['x1'],
                        'y1': line['bottom'],
                    }
                    text_boxes.append((line['text'], bounding_box))

            return text_boxes

def extract_text_within_bbox(pdf_path, bbox):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]  # Assuming you want to extract from the first page
        text_within_bbox = page.within_bbox(bbox).extract_text()
        return text_within_bbox

def extract_tables_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        tables = []
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            table_data = page.extract_tables()
            tables.extend(table_data)
        return tables



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


    # gets bounding box
    if find_keyword_in_pdfs(pdf_file_paths, "Wawanesa"):
        text_bounding_boxes = get_text_bounding_boxes(pdf_path)
        for text, bounding_box in text_bounding_boxes:
            print(f"Text: {text}")
            print(f"Bounding Box: {bounding_box}")
            print()

    # gets table values
        tables = extract_tables_from_pdf(pdf_path)
        for table_number, table in enumerate(tables):
            print(f"Table {table_number + 1}:")
            for row in table:
                print(row)
            print("\n")



    # bb_1 = (46.5600015, 143.63243797500002, 494.03997825, 153.8975676)
    # bb_2 = (46.5600015, 156.67647284999998, 494.03997825, 166.61757285)
    # bb_3 = (46.5600015, 168.67647284999998, 494.03997825, 178.61757285)
    # result = extract_text_within_bbox(pdf_path, bb_1)
    # result_2 = extract_text_within_bbox(pdf_path, bb_2)
    # result_3 = extract_text_within_bbox(pdf_path, bb_3)
    # print(result)
    # print(result_2)
    # print(result_3)