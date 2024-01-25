from pathlib import Path
import pdfplumber

def extract_text_and_bbox(pdf_path):
    result_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            text_boxes = []
            page = pdf.pages[page_number]
            text_lines = page.extract_words()
            for line in text_lines:
                print(line)
                bounding_box = (line['x0'], line['top'], line['x1'], line['bottom'])
                text_boxes.append((line['text'], bounding_box))
            result_dict[page_number + 1] = text_boxes
    return result_dict

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

for pdf_path in pdf_file_paths:
    print(f"\n<========= PDF_FILENAME: {pdf_path} =========>\n")
    text_bounding_boxes = extract_text_and_bbox(pdf_path)
    draw_rectangles(pdf_path, text_bounding_boxes)
    with open( output_dir / f"{Path(pdf_path).stem} Text Coordinates.txt", 'w') as file:
        for page, value in text_bounding_boxes.items():
            file.write(f"\n<========= Page: {page} =========>\n")
            print(f"\n<========= Page: {page} =========>\n")
            for text, box in enumerate(value):
                file.write(f"#{text} : {box}\n")
                print(f"#{text} : {box}")


