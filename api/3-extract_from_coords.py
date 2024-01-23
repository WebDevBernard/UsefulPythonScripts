from pathlib import Path
import pdfplumber

def extract_text_within_bboxes(pdf_path, bboxes):
    with pdfplumber.open(pdf_path) as pdf:
        result = {}
        for page_num, bbox_list in bboxes.items():
            page = pdf.pages[page_num]  # Adjust page index (0-based) to match the dictionary keys
            text_within_bboxes = []
            for bbox in bbox_list:
                text_within_bbox = page.within_bbox(bbox).extract_text()
                text_within_bboxes.append((text_within_bbox, bbox))
            result[page_num] = text_within_bboxes
    return result

def draw_rectangles(pdf_path, rects_list):
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            im = page.to_image(resolution=400)
            if page_number + 1 in rects_list:
                rects = rects_list[page_number + 1]
                y_values = [rect[1] for rect in rects]
                im.draw_rects(y_values)
            if img_capture:
                im.show()

base_dir = Path(__file__).parent.parent
pdf_directory = Path(base_dir / "input")
pdf_files = pdf_directory.glob("*.pdf")
pdf_file_paths = []
output_dir = base_dir / "output"  # name of output folder
output_dir.mkdir(exist_ok=True)

img_capture = {
            "True": "False" 
            }
extract_text_lines_or_words = True # true extract text lines, false extract words

def filename_in_brackets():
    if extract_text_lines_or_words:
        return "(extract_text_lines)"
    else:
        return "(extract_words)" 
# Loop through each PDF file and append the full path to the list

bboxes = {1:[(249.38500023, 31.150000149999983, 304.89500023000005, 41.15000014999998)],
          2:[(48.483000000000004, 30.446999949999963, 68.48100000000001, 39.44699994999996)]
        }
    # print(result_3)
for pdf_file in pdf_files:
    file_path = str(pdf_file)
    pdf_file_paths.append(file_path)

for pdf_path in pdf_file_paths:
    result = extract_text_within_bboxes(pdf_path, bboxes)
    draw_rectangles(pdf_path, result)
    for page, text_list in result.items():
        print(f"\n<========= Page: {page} =========>\n")
        for idx, text in enumerate(text_list, start=1):
            print(f"Page {page}, Bbox {idx}: {text}")

    # Print the result
    # print(result)

