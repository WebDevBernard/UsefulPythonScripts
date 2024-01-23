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
