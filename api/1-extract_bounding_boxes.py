from pathlib import Path
import pdfplumber

def extract_text_and_bbox(pdf_path):
    result_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            text_boxes = []
            page = pdf.pages[page_number]
            # text_lines = page.extract_text_lines()
            text_lines = page.extract_words()
            for line in text_lines:
                bounding_box = (line['x0'], line['top'], line['x1'], line['bottom'])
                text_boxes.append((line['text'], bounding_box))
            result_dict[page_number + 1] = text_boxes
    return result_dict

def extract_text_within_bbox(pdf_path, bbox):
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            text_within_bbox = page.within_bbox(bbox).extract_text()
        return text_within_bbox

def draw_rectangles(pdf_path, rects_list):
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            im = page.to_image(resolution=400)
            if page_number + 1 in rects_list:
                rects = rects_list[page_number + 1]
                y_values = [rect[1] for rect in rects]
                im.draw_rects(y_values)
            im.show()

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
    text_bounding_boxes = extract_text_and_bbox(pdf_path)
    draw_rectangles(pdf_path, text_bounding_boxes)
    for page, value in text_bounding_boxes.items():
        print(f"<========= Page: {page} =========>")
        for text, box in value:
            print(f"text: {text}")
            print(f"\n")
            print(f"bbox: {box}")


