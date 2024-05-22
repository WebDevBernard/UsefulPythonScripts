import re
import fitz
from datetime import datetime
import warnings
import pandas as pd
from collections import defaultdict
from pathlib import Path
from helpers import ff, unique_file_name, find_matching_paths, target_dict

warnings.simplefilter("ignore")


def get_doc_types(doc):
    page = doc[0]
    text_block = page.get_text("text", clip=(409.97900390625, 63.84881591796875, 576.0, 72.82147216796875))
    if "Transaction Timestamp ".casefold() in text_block.casefold():
        return "ICBC"


def search_for_input_dict(doc):
    field_dict = {}
    for page_num in range(len(doc)):
        page = doc[page_num - 1]
        wlist = page.get_text("blocks")
        text_boxes = [list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_num] = [[elem1, elem2] for elem1, elem2 in
                                zip(text_boxes, text_coords)]
    return field_dict


def append_word_to_dict(wlist, field_dict, append_duplicates):
    for words in wlist:
        word = words[4].strip().split("\n")
        if append_duplicates:
            field_dict.append(word)
        if word and word not in field_dict:
            field_dict.append(word)


def search_for_matches(input_dict, type_of_pdf, new_dict):
    field_dict = defaultdict(lambda: defaultdict(list))
    try:
        targets = new_dict[type_of_pdf]
        for pg_num, pg in input_dict.items():
            for i, wlist in enumerate(pg):
                for field_name, target in targets.items():
                    if target is not None:
                        try:
                            # This gets the field using coordinates
                            if target.target_keyword and target.target_coordinates and any(
                                    target.target_keyword in s for s in wlist[0]):
                                coords = tuple(
                                    x + y for x, y in zip(input_dict[pg_num][i][1], target.target_coordinates))
                                coords1 = [pg_num - 1, coords]
                                field_dict[type_of_pdf][field_name].append(coords1)

                            # This gets the field using keyword and indexing
                            elif isinstance(target.target_keyword, str) and any(
                                    target.target_keyword in s for s in wlist[0]):
                                word = input_dict[pg_num][i + target.first_index][0][target.second_index]
                                if target.append_duplicates:
                                    field_dict[type_of_pdf][field_name].append(word)
                                elif target.join_list:
                                    field_dict[type_of_pdf][field_name].append(" ".join(word).split(", "))
                                elif word and word not in field_dict[type_of_pdf][field_name]:
                                    field_dict[type_of_pdf][field_name].append(word)

                            #  This gets the field using keyword (regex) and indexing
                            elif isinstance(target.target_keyword, re.Pattern):
                                word = input_dict[pg_num][i + target.first_index][0][target.second_index]
                                if re.search(target.target_keyword, word) and word not in field_dict[type_of_pdf][
                                    field_name]:
                                    field_dict[type_of_pdf][field_name].append(word)
                        except IndexError:
                            continue
    except KeyError:
        print("This is not an ICBC policy")
    return field_dict


def format_icbc(dict_items, type_of_pdf):
    field_dict = {}
    if type_of_pdf and type_of_pdf == "ICBC":
        if not dict_items["licence_plate"]:
            field_dict["licence_plate"] = "NONLIC"
        for license_plates in dict_items["licence_plate"]:
            for index, license_plate in enumerate(license_plates):
                plate_number = re.sub(re.compile(r"Licence Plate Number "), "", license_plate)
                field_dict["licence_plate"] = plate_number
        for transaction_timestamps in dict_items["transaction_timestamp"]:
            for index, transaction_timestamp in enumerate(transaction_timestamps):
                transaction1 = re.sub(re.compile(r"Transaction Timestamp "), "", transaction_timestamp)
                field_dict["transaction_timestamp"] = transaction1
        for insured_names in dict_items["insured_name"]:
            for index, insured_name in enumerate(insured_names):
                field_dict["insured_name"] = insured_name.rstrip('.')
        for insured_names in dict_items["owner_name"]:
            for index, insured_name in enumerate(insured_names):
                field_dict["insured_name"] = insured_name.rstrip('.')
        for insured_names in dict_items["applicant_name"]:
            for index, insured_name in enumerate(insured_names):
                field_dict["insured_name"] = insured_name.rstrip('.')
    return field_dict


def get_excel_data():
    input_dir = Path(__file__).parent.parent
    excel_path = input_dir / "BM3KXR.xlsx"
    data = {}
    try:
        df_excel = pd.read_excel(excel_path, sheet_name="BM3KXR", header=None)
        data["number_of_pdfs"] = int(df_excel.at[2, 1]) - 1
        data["broker_code"] = df_excel.at[4, 1]
        data["company_name"] = df_excel.at[8, 1]
        data["toggle_customer_copy"] = df_excel.at[13, 1]
    except KeyError:
        return None
    return data


def insert_timestamp(doc, el1):
    page = doc[el1[0]]
    current_time = datetime.today().hour
    date = datetime.today().strftime("%I:%M")
    date_location = tuple(
        x + y for x, y in zip((el1[1]), (0, 1.7, 0, 0))) if current_time < 12 else tuple(
        x + y for x, y in zip((el1[1]), (0, 22.8, 0, 0)))
    page.insert_textbox(date_location, date, align=fitz.TEXT_ALIGN_RIGHT,
                        fontname="spacemo", fontsize=5)


def insert_stamp_position(page, el, data, png_files):
    company_name = data["company_name"]
    broker_code = data["broker_code"]
    text = f"#{broker_code}"
    date = datetime.today().strftime("%b %d, %Y")
    # defines where the stamp gets placed
    fontname = "spacemo"
    fontsize = 6
    fontsize_company_name = 5
    logo_position = tuple(
        x + y for x, y in zip((el[1]), (30, 2, -30, -30)))
    broker_number_above = tuple(
        x + y for x, y in zip((el[1]), (0, 25, 0, 0)))
    broker_number = tuple(
        x + y for x, y in zip((el[1]), (0, 32, 0, 0)))
    broker_number_below = tuple(
        x + y for x, y in zip((el[1]), (0, 40, 0, 0)))
    broker_number_without_logo = tuple(
        x + y for x, y in zip((el[1]), (0, 15, 0, 0)))
    date_without_logo = tuple(
        x + y for x, y in zip((el[1]), (0, 25, 0, 0)))


    if png_files and company_name:
        logo_pic = png_files[0]
        page.insert_image(rect=logo_position, filename=logo_pic, height=0, width=0, keep_proportion=True)
        page.insert_textbox(broker_number_above, company_name, align=fitz.TEXT_ALIGN_CENTER, fontname=fontname,
                            fontsize=fontsize_company_name)
        page.insert_textbox(broker_number, text, align=fitz.TEXT_ALIGN_CENTER, fontname=fontname, fontsize=fontsize)
        page.insert_textbox(broker_number_below, date, align=fitz.TEXT_ALIGN_CENTER, fontname=fontname,
                            fontsize=fontsize)
    if not png_files and company_name:
        page.insert_textbox(broker_number_without_logo, text, align=fitz.TEXT_ALIGN_CENTER, fontname=fontname,
                            fontsize=6)
        page.insert_textbox(date_without_logo, date, align=fitz.TEXT_ALIGN_CENTER, fontname=fontname, fontsize=fontsize)


def write_to_icbc(doc, dict_items):
    input_dir = Path(__file__).parent.parent / "assets"
    png_files = list(input_dir.glob("*.png"))
    data = get_excel_data()
    company_name = data["company_name"]
    if isinstance(company_name, str) and png_files:
        for el in dict_items["ICBC"]["validation_stamp"]:
            page = doc[el[0]]
            insert_stamp_position(page, el, data, png_files)
        if isinstance(dict_items["ICBC"]["time_of_validation"], list):
            for el1 in dict_items["ICBC"]["time_of_validation"]:
                insert_timestamp(doc, el1)
    else:
        for el in dict_items["ICBC"]["validation_stamp"]:
            page = doc[el[0]]
            insert_stamp_position(page, el, data, png_files)
        if isinstance(dict_items["ICBC"]["time_of_validation"], list):
            for el1 in dict_items["ICBC"]["time_of_validation"]:
                insert_timestamp(doc, el1)


def not_customer_copy_page_numbers(pdf):
    pages = []
    with (fitz.open(pdf) as doc):
        top = False
        top_block = doc[0].get_text("text", clip=(230.3990020751953, 36.0, 573.6226196289062, 48.2890625))
        if "Temporary Operation Permit and Owner’s Certificate of Insurance".casefold() in top_block.casefold():
            top = True
        for page_num in range(len(doc)):
            page = doc[page_num]
            text_block = page.get_text("text", clip=(480, 760, 580, 780))
            if not "Customer Copy".casefold() in text_block.casefold():
                pages.append(page_num)
        if top:
            del pages[-1]
    return list(reversed(pages))


def toggle_customer_copy():
    data = get_excel_data()
    toggle = data["toggle_customer_copy"]
    if type(toggle) is str:
        return ""
    else:
        return " (Customer Copy)"


def copy_icbc(number_of_pdfs):
    # input directory sorting
    icbc_input_directory = Path.home() / 'Downloads'
    pdf_files1 = list(icbc_input_directory.glob("*.pdf"))
    pdf_files1 = sorted(pdf_files1, key=lambda file: Path(file).lstat().st_mtime, reverse=True)
    # output directory sorting
    icbc_output_directory = Path(__file__).parent.parent / 'output'
    icbc_output_directory.mkdir(exist_ok=True)
    icbc_output_directory1 = Path(__file__).parent.parent / 'output' / "Customer Copies"
    paths = list(Path(icbc_output_directory).glob("*.pdf"))
    file_names = [path.stem.split()[0] for path in paths]
    toggle = toggle_customer_copy()
    if toggle != " (Customer Copy)":
        icbc_output_directory1.mkdir(exist_ok=True)
    for pdf in pdf_files1[:number_of_pdfs + 1]:
        with (fitz.open(pdf) as doc):
            doc_type = get_doc_types(doc)
            input_dict = search_for_input_dict(doc)
            dict_items = search_for_matches(input_dict, doc_type, target_dict)
            formatted_dict = format_icbc(ff(dict_items[doc_type]), doc_type)
            try:
                df = pd.DataFrame([formatted_dict])
                print(df)
            except KeyError:
                continue
            if doc_type and doc_type == "ICBC":
                icbc_file_name = f"{df['licence_plate'].at[0]}{toggle}.pdf" if df["licence_plate"].at[
                                        0] != "NONLIC" else f"{df["insured_name"].at[0].title()}{toggle}.pdf"
                icbc_file_name1 = f"{df['licence_plate'].at[0]} (Customer Copy).pdf" if df["licence_plate"].at[
                                        0] != "NONLIC" else f"{df["insured_name"].at[0].title()} (Customer Copy).pdf"

                icbc_output_path = icbc_output_directory / icbc_file_name
                icbc_output_path1 = icbc_output_directory1 / icbc_file_name1

                target_filename = Path(icbc_file_name).stem.split()[0]
                matching_transaction_ids = []
                if target_filename in file_names:
                    # write_to_icbc(doc, dict_items)
                    # doc.delete_pages(not_customer_copy_page_numbers(pdf))
                    # doc.save(unique_file_name(icbc_output_path))
                    matching_paths = find_matching_paths(target_filename, paths)
                    for path_name in matching_paths:
                        with fitz.open(path_name) as doc1:
                            target_transaction_id = doc1[0].get_text("text", clip=(
                                409.97900390625, 63.84881591796875, 576.0, 72.82147216796875))
                            match = int(re.match(re.compile(r'.*?(\d+)'), target_transaction_id).group(1))
                            matching_transaction_ids.append(match)
                if int(df["transaction_timestamp"].at[0]) not in matching_transaction_ids:
                    write_to_icbc(doc, dict_items)
                    if toggle == " (Customer Copy)":
                        doc.delete_pages(not_customer_copy_page_numbers(pdf))
                        doc.save(unique_file_name(icbc_output_path), garbage=4, deflate=True)
                    else:
                        doc.save(unique_file_name(icbc_output_path), garbage=4, deflate=True)
                        doc.delete_pages(not_customer_copy_page_numbers(pdf))
                        doc.save(unique_file_name(icbc_output_path1), garbage=4, deflate=True)


def main():
    data = get_excel_data()
    copy_icbc(data["number_of_pdfs"])


if __name__ == "__main__":
    main()
