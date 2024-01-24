from pathlib import Path
import pdfplumber

def extract_text_within_bbox(pdf_path, page_number, bbox):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        cropped_page = page.crop(bbox)
        text = cropped_page.extract_text()
    return text

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

def draw_rectangles(pdf_path, page_number, rects_list):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        page.to_image(resolution=400).draw_rects([rects_list]).show()

pg = 3
coords = (36.4375, 67.22000049999991, 170.99999650000004, 94.84000049999992)

# Extract text from page 1 using coordinates from the dictionary
for pdf_path in pdf_file_paths:
    draw_rectangles(pdf_path, pg, coords)
    extracted_text = extract_text_within_bbox(pdf_path, pg, coords)
    print(extracted_text)
