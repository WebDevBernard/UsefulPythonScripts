from pathlib import Path
import pdfplumber
import pandas as pd

# Table Strategy for identifying table, using h-lines or h-text
ts = {
    "vertical_strategy": "explicit",
    "explicit_vertical_lines": [23.64, 50.64, 106.91, 159.36, 217.94, 341.64, 392.67, 430.92, 637.23, 677.32],
    "horizontal_strategy": "text",
    "snap_y_tolerance": 6,
}

# ts = {
#     "vertical_strategy": "lines",
#     "horizontal_strategy": "lines",
# }

def extract_tables_from_pdf(pdf_path):
    dictionary = {}
    text_boxes = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            page = page.crop(bbox=(23.64, 117.61, 677.32, 586.69))
            # Debug screen to show where table is drawn
            # page.to_image(resolution=400).debug_tablefinder(ts).show()

            # Get all the row coordinates
            row_coords = []
            find_tables = page.find_tables(ts)
            for table in find_tables:
                for row in table.rows:
                    row_coords.append(row.bbox)
            extract_tables = page.extract_tables(ts)

            # Get all table texts
            table_texts = []
            for table in extract_tables:
                table_texts.extend(table)
            dictionary[page_number + 1] = dictionary[page_number + 1] = [[elem1, elem2] for elem1, elem2 in zip(table_texts, row_coords)]

    return dictionary

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
    pages = extract_tables_from_pdf(pdf_path)
    with open(output_dir / f"{Path(pdf_path).stem} extract_table.txt", 'w') as file:
        loc_cols = ["insurer", "time_entered", "time_made", "policynum", "csrcode", "ccode", "name",
                    "renewal"]
        rows = [item[0] for sublist in pages.values() for item in sublist]
        df = pd.DataFrame(rows, columns=loc_cols)
        df = df.dropna()
        print(df)
