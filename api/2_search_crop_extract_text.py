from pathlib import Path
import pdfplumber

search = "SMOKE DETECTOR RATE APPLIED"

def extract_tables_from_pdf(pdf_path):
    dictionary = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            text_boxes = []
            description = page.search(search)
            for line in description:
                bounding_box = (line['x0'], line['top'], line['x1'], line['bottom'])
                text_boxes.append([line['text'], bounding_box])
            dictionary[page_number + 1] = text_boxes
    return dictionary

def extract_text_within_bbox(pdf_path, page_number, bbox):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        cropped_page = page.crop(bbox)
        page.to_image(resolution=400).draw_rect(bbox).show()
        text = cropped_page.extract_text()
    return text

def draw_rectangles(pdf_path, dictionary):
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            if page_number + 1 in dictionary:
                dict_list = dictionary[page_number + 1]
                pdf.pages[page_number].to_image(resolution=400).draw_rects([x[1] for x in dict_list]).show()

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

pg = 1
coords = list([[110.64081612470625, 101.41569021025873, 179.91378549927597, 109.45572843850499], [23.64, 372.49708401291565, 677.32, 386.05949620041565]][0])

for pdf_path in pdf_file_paths:
    found_search = extract_tables_from_pdf(pdf_path)
    draw_rectangles(pdf_path, found_search)
    # extracted_text = extract_text_within_bbox(pdf_path, pg, coords)
    print(found_search)

