from pathlib import Path
import pdfplumber

ts = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "text"
}


def filter_dict(dictionary):
        return {key: [sublist for sublist in ([item for item in sublist if item is not None and item != ''] for sublist in sublists) if sublist] for key, sublists in dictionary.items()}

def extract_tables_from_pdf(pdf_path):
    result_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number in range(len(pdf.pages)):
            tables = []
            page = pdf.pages[page_number]
            extract_tables = page.extract_tables(ts)
            for table in extract_tables:
                tables.extend(table)
            result_dict[page_number + 1] = tables
    return filter_dict(result_dict)

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
        for page, value in pages.items():
            file.write(f"\n<========= Page: {page} =========>\n")
            print(f"\n<========= Page: {page} =========>\n")
            for text, box in enumerate(value):
                file.write(f"#{text} : {box}\n")
                print(f"#{text} : {box}")
