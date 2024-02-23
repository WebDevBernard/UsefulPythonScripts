import fitz
import pandas as pd
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader, PdfWriter
from debugging import base_dir
from cleaning_functions import (search_first_page, search_for_matches,
                                search_for_input_dict, get_broker_copy_pages, format_policy, create_pandas_df)
from coordinates import doc_type, keyword, target_dict, filename
from formatting_functions import unique_file_name
# Loop through each PDF file and append the full path to the list

input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

def write_to_new_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["risk_type"].title()}.docx"
    doc.save(unique_file_name(output_path))


def renewal_letter(excel_path):
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            print(f"\n<==========================>\n\nFilename is: {Path(pdf).stem}{Path(pdf).suffix} ")
            type_of_pdf = search_first_page(doc, doc_type)
            pg_list = get_broker_copy_pages(doc, type_of_pdf, keyword)
            input_dict = search_for_input_dict(doc, pg_list)
            dict_items = search_for_matches(doc, input_dict, type_of_pdf, target_dict)
            cleaned_data = format_policy(dict_items, type_of_pdf)
            try:
                df = create_pandas_df(cleaned_data)
            except KeyError:
                continue
            df["broker_name"] = pd.read_excel(excel_path, sheet_name=0, header=None).at[8, 1]
            df["mods"] = pd.read_excel(excel_path, sheet_name=0, header=None).at[4, 1]
            print(df)
            for rows in df.to_dict(orient="records"):
                write_to_new_docx(filename["RL"], rows)
            print(f"\n<==========================>\n")

def renewal_letter_manual(excel_data):
    df = pd.DataFrame([excel_data])
    df["today"] = datetime.today().strftime("%B %d, %Y")
    df["mailing_address"] = df[["address_line_one", "address_line_two"]].astype(str).apply(
        lambda x: ', '.join(x[:-1]) + " " + x[-1:], axis=1)
    df["risk_address_1"] = df["risk_address_1"].fillna(df["mailing_address"])
    df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
    print(df)
    for rows in df.to_dict(orient="records"):
        write_to_new_docx(filename["RL"], rows)

# Used for filling docx
def write_to_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["type"].title()}.docx"
    doc.save(unique_file_name(output_path))

# Used for fillable pdf's
def write_to_pdf(pdf, dictionary, rows):
    pdf_path = (base_dir / "templates" / pdf)
    output_path = base_dir / "output" / f"{rows["named_insured"]} {rows["risk_type"].title()}.pdf"
    output_path.parent.mkdir(exist_ok=True)
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        writer.add_page(page)
        writer.updatePageFormFieldValues(page, dictionary)
    with open(unique_file_name(output_path), "wb") as output_stream:
        writer.write(output_stream)