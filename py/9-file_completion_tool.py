import fitz
import pandas as pd
from datetime import datetime
from pathlib import Path
from debugging import base_dir
from file_writing_functions import (search_first_page, search_for_matches, write_to_new_docx,
                                    search_for_input_dict, get_broker_copy_pages, format_policy, create_pandas_df)
from coordinates import doc_type, keyword, target_dict, filename
from sort_renewal_list import sort_renewal_list
# Loop through each PDF file and append the full path to the list
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")
excel_path = base_dir / "input.xlsx"  # name of Excel

def renewal_letter():
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
            df["broker_name"] = pd.read_excel(excel_path, sheet_name=0, header=None).at[10, 1]
            print(df)
            for rows in df.to_dict(orient="records"):
                write_to_new_docx(filename["RL"], rows)
            print(f"\n<==========================>\n")

def get_excel_data():
    data = {}
    try:
        df = pd.read_excel(excel_path, sheet_name=0, header=None)
        data["broker_name"] = df.at[10, 1]
        data["risk_type"] = df.at[15, 1]
        data["named_insured"] = df.at[18, 1]
        data["insurer"] = df.at[19, 1]
        data["policy_number"] = df.at[20, 1]
        data["effective_date"] = str(df.at[21, 1])
        data["address_line_one"] = df.at[23, 1]
        data["address_line_two"] = df.at[24, 1]
        data["address_line_three"] = df.at[25, 1]
        data["risk_address_1"] = df.at[27, 1]
    except KeyError:
        return None
    return data

def renewal_letter_manual():
    df = pd.DataFrame([get_excel_data()])
    df["today"] = datetime.today().strftime("%B %d, %Y")
    df["mailing_address"] = df[["address_line_one", "address_line_two"]].astype(str).apply(
        lambda x: ', '.join(x[:-1]) + " " + x[-1:], axis=1)
    df["risk_address_1"] = df["risk_address_1"].fillna(df["mailing_address"])
    print(df)
    for rows in df.to_dict(orient="records"):
        write_to_new_docx(filename["RL"], rows)

if __name__ == "__main__":
    df_excel = pd.read_excel(excel_path, sheet_name=0, header=None)
    task = df_excel.at[6, 1]
    if task == "Auto Generate Renewal Letter From Policy":
        renewal_letter()
    elif task == "Manually Generate Renewal Letter from Data Entry":
        renewal_letter_manual()
    elif task == "Sort Renewal List from Excel File":
        sort_renewal_list()
