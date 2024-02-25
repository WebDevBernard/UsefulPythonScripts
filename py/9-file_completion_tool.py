import pandas as pd
import fitz
import os
from pathlib import Path
from docxtpl import DocxTemplate
from formatting_functions import unique_file_name

base_dir = Path(__file__).parent.parent

def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path

def write_to_new_docx(docx, rows):
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    template_path = base_dir / "templates" / docx
    doc = DocxTemplate(template_path)
    doc.render(rows)
    output_path = output_dir / f"{rows["named_insured"]} {rows["risk_type"].title()}.docx"
    doc.save(unique_file_name(output_path))


def get_excel_path():
    xlsx_files = Path(input_dir).glob("*.xlsx")
    xls_files = Path(input_dir).glob("*.xls")
    files = list(xlsx_files) + list(xls_files)
    return list(files)[0]


excel_path = get_excel_path()


def make_df_from_excel():
    if Path(excel_path).suffix == ".XLS":
        return pd.read_excel(excel_path, engine="xlrd")
    elif Path(excel_path).suffix == ".XLSX":
        return pd.read_excel(excel_path, engine="openpyxl")


def sort_renewal_list():
    try:
        output_path = base_dir / "output" / f"{Path(excel_path).stem}.xlsx"
        df = make_df_from_excel()
        column_list = ["policynum", "ccode", "name", "pcode", "csrcode", "insurer", "buscode", "renewal", "Pulled",
                       "D/L"]
        df = df.reindex(columns=column_list)
        df = df.drop_duplicates(subset=["policynum"], keep=False)
        df.sort_values(["insurer", "renewal", "name"], ascending=[True, True, True], inplace=True)
        list_with_spaces = []
        for x, y in df.groupby('insurer', sort=False):
            list_with_spaces.append(y)
            list_with_spaces.append(pd.DataFrame([[float('NaN')] * len(y.columns)], columns=y.columns))
        df = pd.concat(list_with_spaces, ignore_index=True).iloc[:-1]
        print(df)
        if not os.path.isfile(output_path):
            writer = pd.ExcelWriter(output_path, engine="openpyxl")
        else:
            writer = pd.ExcelWriter(output_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        writer.close()
    except TypeError:
        return


input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")
excel_path = base_dir / "input.xlsx"  # name of Excel

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


def get_excel_data(excel_path):
    data = {}
    try:
        df = pd.read_excel(excel_path, sheet_name=0, header=None)
        data["broker_name"] = df.at[8, 1]
        data["risk_type"] = df.at[13, 1]
        data["named_insured"] = df.at[15, 1]
        data["insurer"] = df.at[16, 1]
        data["policy_number"] = df.at[17, 1]
        data["effective_date"] = df.at[18, 1]
        data["address_line_one"] = df.at[20, 1]
        data["address_line_two"] = df.at[21, 1]
        data["address_line_three"] = df.at[22, 1]
        data["risk_address_1"] = df.at[24, 1]
    except KeyError:
        return None
    return data

excel_data = get_excel_data(excel_path)

if __name__ == "__main__":
    df_excel = pd.read_excel(excel_path, sheet_name=0, header=None)
    task = df_excel.at[2, 1]
    if task == "Auto Generate Renewal Letter":
        renewal_letter(excel_path)
    elif task == "Manual Renewal Letter":
        renewal_letter_manual(excel_data)
    elif task == "Sort Renewal List":
        sort_renewal_list()
