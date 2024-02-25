import fitz
import re
from pathlib import Path
from collections import namedtuple, defaultdict
from Helpers import target_dict, and_regex, address_regex, date_regex, dollar_regex, \
    postal_code_regex, ff, find_index, join_and_format_names, address_one_title_case, address_two_title_case


def get_doc_types(doc):
    DocType = namedtuple("Insurer", "pdf_name keyword coordinates", defaults=None)
    doc_types = [
        DocType("Aviva", "Company", (171.36000061035156, 744.800048828125, 204.39999389648438, 752.7999877929688)),
        DocType("Family", "Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)),
        DocType("Intact", "BROKER COPY", (250, 764.2749633789062, 360, 773.8930053710938)),
        DocType("Wawanesa", "BROKER OFFICE", (36.0, 102.42981719970703, 353.2679443359375, 111.36731719970703))
    ]
    for doc_type in doc_types:
        for page_index in range(len(doc)):
            page = doc[page_index]
            text_block = page.get_text("text", clip=doc_type.coordinates)
            if doc_type.keyword.casefold() in text_block.casefold():
                return doc_type.pdf_name


def get_content_pages(doc, pdf_name):
    ContentPages = namedtuple("Insurer", "pdf_name keyword coordinates", defaults=None)
    content_pages = [
        ContentPages("Aviva", "CANCELLATION OF THE POLICY", -1),
        ContentPages("Intact", "BROKER COPY", (250, 764.2749633789062, 360, 773.8930053710938)),
        ContentPages("Wawanesa", re.compile(r"\w{3}\s\d{2},\s\d{4}"),
                     (36.0, 762.829833984375, 576.001220703125, 778.6453857421875))
    ]
    pg_list = []
    for content_page in content_pages:
        keyword = content_page.keyword
        coordinates = content_page.coordinates
        for page_index in range(len(doc)):
            page = doc[page_index]
            # if the type of pdf found in get_doc_types matches the pdf_name
            if content_page.pdf_name == pdf_name:

                # Finds which page to stop on with coordinates being the starting range of pages
                if isinstance(keyword, str) and isinstance(coordinates, int):
                    if page.search_for(keyword):
                        for page_num in range(page_index + coordinates, page_index + 1):
                            pg_list.append(page_num + 1)

                # Finds which page contain a string match to a clipped area on the page
                elif isinstance(keyword, str) and isinstance(coordinates, tuple):
                    if page.search_for(keyword, clip=coordinates):
                        pg_list.append(page_index + 1)

                # Finds which page contain a regex match to a clipped area on the page
                elif isinstance(keyword, re.Pattern) and isinstance(coordinates, tuple):
                    for word_list in page.get_text("blocks", clip=coordinates):
                        if keyword.search(word_list[4]):
                            pg_list.append(page_index + 1)

    # If no search condition, return count of all pages
    if not pg_list:
        for page_num in range(doc.page_count):
            pg_list.append(page_num + 1)
    return pg_list


def search_for_input_dict(doc, pg_list):
    field_dict = {}
    for page_num in pg_list:
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


def search_for_matches(doc, input_dict, type_of_pdf, target_dict):
    field_dict = defaultdict(lambda: defaultdict(list))
    try:
        targets = target_dict[type_of_pdf]
        for pg_num, pg in input_dict.items():
            page = doc[pg_num - 1]
            for i, wlist in enumerate(pg):
                for field_name, target in targets.items():
                    if target is not None:
                        try:
                            # This gets the name_and_address fields
                            if target.target_coordinates and target.target_keyword is None:
                                tuple_list = doc[0].get_text("blocks", clip=target.target_coordinates)
                                append_word_to_dict(tuple_list, field_dict[type_of_pdf][field_name],
                                                    target.append_duplicates)

                            # This gets the field using coordinates
                            elif target.target_keyword and target.target_coordinates and any(
                                    target.target_keyword in s for s in wlist[0]):
                                coords = tuple(
                                    x + y for x, y in zip(input_dict[pg_num][i][1], target.target_coordinates))
                                word_list = page.get_text("blocks", clip=coords)
                                append_word_to_dict(word_list, field_dict[type_of_pdf][field_name],
                                                    target.append_duplicates)

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
        print("This pdf dictionary has not been created yet")
    return field_dict


def format_named_insured(field_dict, dict_items, type_of_pdf):
    for name_and_address in dict_items["name_and_address"]:
        address_index = find_index(address_regex, name_and_address)
        if type_of_pdf == "Aviva" or type_of_pdf == "Family" or type_of_pdf == "Wawanesa":
            names = re.sub(and_regex, "", ", ".join(name_and_address[:address_index]))
            field_dict["named_insured"] = join_and_format_names(names.split(", "))
        if type_of_pdf == "Intact":
            names = [i.split(' & ') for i in name_and_address[:address_index]]
            join_same_last_names = [" ".join(reversed(i.split(", "))) if ", " in i else
                                    (i + " " + names[0][0]).split(", ")[0] for i in names[0]]
            field_dict["named_insured"] = join_and_format_names(join_same_last_names)
    return field_dict


def format_insurer_name(field_dict, type_of_pdf):
    field_dict["insurer"] = type_of_pdf
    return field_dict


def format_mailing_address(field_dict, dict_items):
    for name_and_address in dict_items["name_and_address"]:
        pc_index = find_index(postal_code_regex, name_and_address)
        address_index = find_index(address_regex, name_and_address)
        city_province_p_code = " ".join(name_and_address[address_index + 1:pc_index + 1])
        field_dict["address_line_one"] = address_one_title_case(" ".join(name_and_address[address_index:pc_index - 1]))
        field_dict["address_line_two"] = address_two_title_case(
            re.sub(re.compile(r"Canada,"), "", re.sub(postal_code_regex, "", city_province_p_code)))
        field_dict["address_line_three"] = re.search(postal_code_regex, city_province_p_code).group().title()
    return field_dict


def format_policy_number(field_dict, dict_items):
    field_dict["policy_number"] = dict_items["policy_number"][0][0]
    return field_dict


def format_effective_date(field_dict, dict_items):
    field_dict["effective_date"] = re.search(date_regex, dict_items["effective_date"][0][0]).group()
    return field_dict


def sum_dollar_amounts(amounts):
    clean_amount_str = [a.replace("$", "").replace(",", "").replace(" 00", "").replace(".00", "") for a in amounts[0]]
    total = sum(int(c) for c in clean_amount_str)
    return total


def format_premium_amount(field_dict, dict_items):
    field_dict["premium_amount"] = '${:,.2f}'.format(sum_dollar_amounts(dict_items["premium_amount"]))
    return field_dict


def format_additional_coverage(field_dict, dict_items, type_of_pdf):
    if (type_of_pdf == "Family" and dict_items["earthquake_coverage"] and
            re.search(dollar_regex, dict_items["earthquake_coverage"][0][0])):
        field_dict["earthquake_coverage"] = True
    if type_of_pdf == "Aviva" or type_of_pdf == "Intact" or type_of_pdf == "Wawanesa" and dict_items[
        "earthquake_coverage"]:
        field_dict["earthquake_coverage"] = True
    if type_of_pdf == "Intact" and dict_items["ground_water"]:
        field_dict["ground_water"] = True
    if type_of_pdf == "Wawanesa" and dict_items["tenant_vandalism"]:
        field_dict["tenant_vandalism"] = True
    if dict_items["overland_water"]:
        field_dict["overland_water"] = True
    if dict_items["service_line"]:
        field_dict["service_line"] = True


def find_nested_match(regex, nested_list):
    matched_lists = []
    for item in nested_list:
        if re.search(regex, item) and item not in matched_lists:
            matched_lists.append(item)
    return matched_lists


def find_risk_addresses(risk_addresses):
    matched = []
    for index, risk_address in enumerate(risk_addresses):
        if re.search(postal_code_regex, " ".join(risk_address)) and risk_address not in matched:
            matched.append(risk_address)
    return matched


def format_risk_address(field_dict, dict_items):
    risk_addresses = find_risk_addresses(dict_items["risk_address"])
    for index, risk_address in enumerate(risk_addresses):
        field_dict[f"risk_address_{index + 1}"] = address_two_title_case(
            re.sub(postal_code_regex, "", risk_address[0]).rstrip(", "))
    return field_dict

def flatten(xss):
    return [item for sublist in xss for item in sublist]

def format_form_type(field_dict, dict_items, type_of_pdf):
    for form_types in dict_items["form_type"]:
        for index, form_type in enumerate(form_types):
            if "comprehensive".casefold() in form_type.casefold():
                field_dict[f"form_type_{index + 1}"] = "Comprehensive Form"
            if "broad".casefold() in form_type.casefold():
                field_dict[f"form_type_{index + 1}"] = "Broad Form"
            if "basic".casefold() in form_type.casefold():
                field_dict[f"form_type_{index + 1}"] = "Basic Form"
            if "fire & extended".casefold() in form_type.casefold():
                field_dict[f"form_type_{index + 1}"] = "Fire + EC"
            if type_of_pdf == "Family":
                if "included".casefold() in form_type.casefold():
                    field_dict[f"form_type_{index + 1}"] = "Comprehensive Form"
                else:
                    field_dict[f"form_type_{index + 1}"] = "Specified Perils"
    return field_dict

def format_risk_type(field_dict, dict_items, type_of_pdf):
    for risk_types in dict_items["risk_type"]:
        for index, risk_type in enumerate(risk_types):
            if "seasonal".casefold() in risk_type.casefold():
                field_dict["seasonal"] = True
            if "home".casefold() in risk_type.casefold():
                field_dict[f"risk_type_{index+1}"] = "home"
            elif "rented dwelling".casefold() in risk_type.casefold():
                field_dict[f"risk_type_{index+1}"] = "rented_dwelling"
            elif "revenue".casefold() in risk_type.casefold():
                field_dict[f"risk_type_{index+1}"] = "rented_dwelling"
            elif "rental".casefold() in risk_type.casefold():
                field_dict[f"risk_type_{index+1}"] = "rented_condo"
            elif "tenant".casefold() in risk_type.casefold():
                field_dict[f"risk_type_{index+1}"] = "tenant"
            elif type_of_pdf == "Aviva" and "condominium".casefold() in risk_type.casefold():
                    field_dict[f"risk_type_{index+1}"] = "condo"
            elif type_of_pdf == "Family" and "condo".casefold() in risk_type.casefold():
                    field_dict[f"risk_type_{index+1}"] = "condo"
            elif type_of_pdf == "Intact" and "condominium ".casefold() in risk_type.casefold():
                    field_dict[f"risk_type_{index+1}"] = "condo"
            elif type_of_pdf == "Wawanesa" and "Condominium" in risk_type:
                    field_dict[f"risk_type_{index+1}"] = "condo"
    return field_dict

def match_keyword(dict_of_keywords, keyword):
    return dict_of_keywords.get(keyword, None)

def format_number_families(field_dict, dict_items, type_of_pdf):
    dict_of_keywords = {
        "One": 1,
        "Two": 2,
        "Three": 3,
        "1": 1,
        "2": 2,
        "3": 3,
        "001 Additional Family": 2,
        "002 Additional Family": 3,
    }
    if not dict_items["number_of_families"]:
        field_dict["number_of_families_1"] = match_keyword(dict_of_keywords, "1")
    for families in dict_items["number_of_families"]:
        for index, number_of_families in enumerate(families):
            if type_of_pdf == "Aviva":
                field_dict[f"number_of_families_{index + 1}"] = match_keyword(dict_of_keywords, number_of_families)
            if type_of_pdf == "Family":
                match = re.search(r"\b(\d+)\b", number_of_families)
                if match:
                    number = str(int(match.group(1)) + 1)
                    field_dict[f"number_of_families_{index+1}"] = match_keyword(dict_of_keywords, number)
            if type_of_pdf == "Intact" or type_of_pdf == "Wawanesa":
                field_dict[f"number_of_families_{index+1}"] = match_keyword(dict_of_keywords, number_of_families)
    return field_dict

# def format_condo_deductible(field_dict, dict_items, type_of_pdf):
#     try:
#         if isinstance(dict_items["condo_deductible"], str):
#             field_dict["condo_deductible_1"] = return_match_only(dollar_regex,
#                                                                  dict_items["condo_deductible"])
#             if type_of_pdf == "Aviva" and dict_items["condo_deductible"]:
#                 field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex, dict_items["condo_deductible"])
#         else:
#             for index, condo_deductible in enumerate(dict_items["condo_deductible"]):
#                 field_dict[f"condo_deductible_{index + 1}"] = return_match_only(dollar_regex, condo_deductible)
#                 if type_of_pdf == "Aviva" and dict_items["condo_deductible"]:
#                     field_dict["condo_earthquake_deductible_1"] = return_match_only(dollar_regex, condo_deductible)
#         if type_of_pdf == "Family" and dict_items["condo_deductible"]:
#             field_dict["condo_deductible_1"] = return_match_only(dollar_regex, dict_items["condo_deductible"][0])
#     except TypeError:
#          return
#     return field_dict


def format_condo_earthquake_deductible(field_dict, dict_items, type_of_pdf):
    for deductibles in dict_items["condo_earthquake_deductible"]:
        for index, condo_earthquake_deductible in enumerate(deductibles):
            if type_of_pdf == "Family" and dict_items["condo_deductible"]:

                field_dict["condo_earthquake_deductible_1"] = re.search(dollar_regex,
                                                                                dict_items["condo_deductible"][1]).group
            elif type_of_pdf == "Intact" and dict_items["condo_earthquake_deductible"]:
                field_dict["condo_earthquake_deductible_1"] = "$25,000"
            elif type_of_pdf == "Intact" and not dict_items["condo_earthquake_deductible"]:
                field_dict["condo_earthquake_deductible_1"] = "$2,500"
            else:
                field_dict[f"condo_earthquake_deductible_{index + 1}"] = re.search(dollar_regex,
                                                                                       condo_earthquake_deductible).group
    return field_dict


def format_policy(flattened_dict, type_of_pdf):
    field_dict = {}
    if type_of_pdf:
        format_named_insured(field_dict, flattened_dict, type_of_pdf)
        format_insurer_name(field_dict, type_of_pdf)
        format_mailing_address(field_dict, flattened_dict)
        format_policy_number(field_dict, flattened_dict)
        format_effective_date(field_dict, flattened_dict)
        format_premium_amount(field_dict, flattened_dict)
        format_additional_coverage(field_dict, flattened_dict, type_of_pdf)
        format_risk_address(field_dict, flattened_dict)
        format_form_type(field_dict, flattened_dict, type_of_pdf)
        format_risk_type(field_dict, flattened_dict, type_of_pdf)
        format_number_families(field_dict, flattened_dict, type_of_pdf)
        # format_condo_deductible(field_dict, flattened_dict, type_of_pdf)
        format_condo_earthquake_deductible(field_dict, flattened_dict, type_of_pdf)
    return field_dict

# 6 append to Pandas Dataframe:
# def create_pandas_df(data_dict):
#     df = pd.DataFrame([data_dict])
#     df["today"] = datetime.today().strftime("%B %d, %Y")
#     df["effective_date"] = pd.to_datetime(df["effective_date"]).dt.strftime("%B %d, %Y")
#     expiry_date = pd.to_datetime(df["effective_date"]) + pd.offsets.DateOffset(years=1)
#     df["expiry_date"] = expiry_date.dt.strftime("%B %d, %Y")
#     return df
#
# def write_to_pdf(pdf, dictionary, rows):
#     pdf_path = (base_dir / "templates" / pdf)
#     output_path = base_dir / "output" / f"{rows["named_insured"]} {rows["risk_type"].title()}.pdf"
#     output_path.parent.mkdir(exist_ok=True)
#     reader = PdfReader(pdf_path)
#     writer = PdfWriter()
#     for page_num in range(len(reader.pages)):
#         page = reader.pages[page_num]
#         writer.add_page(page)
#         writer.updatePageFormFieldValues(page, dictionary)
#     with open(unique_file_name(output_path), "wb") as output_stream:
#         writer.write(output_stream)


def main():
    for pdf_file in pdf_files:
        with fitz.open(pdf_file) as doc:
            print(f"\n<==========================>\n\nFilename is: {Path(pdf_file).stem}{Path(pdf_file).suffix} ")

            # Look through Pdf Pages to find the matching pdf
            doc_type = get_doc_types(doc)
            print(f"This is a {doc_type} policy.")

            # Look for pages with the content for each field
            pg_list = get_content_pages(doc, doc_type)
            print(
                f"\nBroker copies / coverage summary / certificate of property insurance located on pages:\n{pg_list}\n")

            # Extract the dictionary containing for all the strings and corresponding coordinates:
            input_dict = search_for_input_dict(doc, pg_list)
            dict_items = search_for_matches(doc, input_dict, doc_type, target_dict)
            # print(ff(dict_items[doc_type]))
            print(format_policy(ff(dict_items[doc_type]), doc_type))
            print(f"\n<==========================>\n")


base_dir = Path(__file__).parent.parent
input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")

if __name__ == "__main__":
    main()
