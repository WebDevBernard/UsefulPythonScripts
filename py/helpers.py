import os
import re
from collections import namedtuple
from pathlib import Path

postal_code_regex = re.compile(
    r"([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d)$"
)
dollar_regex = re.compile(r"\$([\d,]+)")
date_regex = re.compile(
    r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}|(Jan(uary)?|Feb(ruary)?"
    r"|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?"
    r"|Dec(ember)?)\s+\d{1,2},\s+\d{4}"
)
address_regex = re.compile(
    r"(?!.*\bltd\.)((po box)|(unit)|\d+\s+)", flags=re.IGNORECASE
)
and_regex = re.compile(r"&|\b(and)\b", flags=re.IGNORECASE)

form_fields = [
    "insurer",
    "name_and_address",
    "policy_number",
    "effective_date",
    "risk_address",
    "form_type",
    "risk_type",
    "number_of_families",
    "number_of_units",
    "premium_amount",
    "condo_deductible",
    "condo_earthquake_deductible",
    "earthquake_coverage",
    "overland_water",
    "ground_water",
    "tenant_vandalism",
    "service_line",
    "licence_plate",
    "transaction_type",
    "name_code",
    "transaction_timestamp",
    "insured_name",
    "owner_name",
    "applicant_name",
    "validation_stamp",
    "customer_copy",
    "time_of_validation",
]
FormFields = namedtuple("FormFields", form_fields, defaults=(None,) * len(form_fields))
target_fields = [
    "target_keyword",
    "first_index",
    "second_index",
    "target_coordinates",
    "append_duplicates",
    "join_list",
]
TargetFields = namedtuple(
    "TargetFields", target_fields, defaults=(None,) * len(target_fields)
)
remaining_index = slice(
    *map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(":"))
)

DocType = namedtuple("DocType", "pdf_name keyword coordinates", defaults=None)
doc_types = [
    DocType(
        "Aviva",
        "Company",
        (171.36000061035156, 744.800048828125, 204.39999389648438, 752.7999877929688),
    ),
    DocType(
        "Aviva",
        "Aviva",
        (183.83999633789062, 728.4000244140625, 197.9759979248047, 734.4000244140625),
    ),
    DocType(
        "Family",
        "Agent",
        (25.70400047302246, 36.369102478027344, 51.03702926635742, 45.4451789855957),
    ),
    DocType("Intact", "KMJ", (475, 15, 576, 69)),
    DocType(
        "Intact",
        "Version",
        (36, 633.0250244140625, 576.2601318359375, 767.948974609375),
    ),
    DocType(
        "Intact", "Intact", (36, 633.0250244140625, 576.2601318359375, 767.948974609375)
    ),
    DocType(
        "Wawanesa",
        "BROKER OFFICE",
        (36.0, 102.42981719970703, 353.2679443359375, 111.36731719970703),
    ),
    DocType(
        "ICBC",
        "Transaction Timestamp ",
        (409.97900390625, 63.84881591796875, 576.0, 72.82147216796875),
    ),
]

ContentPages = namedtuple("ContentPages", "pdf_name keyword coordinates", defaults=None)
content_pages = [
    ContentPages("Aviva", "CANCELLATION OF THE POLICY", -1),
    ContentPages(
        "Intact", "BROKER COPY", (250, 764.2749633789062, 360, 773.8930053710938)
    ),
    ContentPages(
        "Wawanesa",
        re.compile(r"\w{3}\s\d{2},\s\d{4}"),
        (36.0, 762.829833984375, 576.001220703125, 778.6453857421875),
    ),
    ContentPages(
        "ICBC",
        "ICBC Copy",
        (36.0, 761.039794921875, 560.1397094726562, 769.977294921875),
    ),
]
target_dict = {
    "Aviva": FormFields(
        name_and_address=TargetFields(
            target_coordinates=(80.4000015258789, 202.239990234375, 250, 280)
        ),
        policy_number=TargetFields(
            target_keyword="Policy Number",
            target_coordinates=(
                267.11999893188477,
                10.15997314453125,
                -202.8189697265625,
                9.16009521484375,
            ),
        ),
        effective_date=TargetFields(
            target_keyword="Policy Effective From:", first_index=0, second_index=1
        ),
        risk_address=TargetFields(
            target_keyword="FORM",
            first_index=-1,
            second_index=remaining_index,
            join_list=True,
        ),
        form_type=TargetFields(
            target_keyword=re.compile(r".* - .*form$", flags=re.IGNORECASE),
            first_index=0,
            second_index=0,
            append_duplicates=True,
        ),
        risk_type=TargetFields(
            target_keyword=re.compile(r".* - .*form$", flags=re.IGNORECASE),
            first_index=0,
            second_index=0,
            append_duplicates=True,
        ),
        number_of_families=TargetFields(
            target_keyword="Extended Liability", first_index=1, second_index=0
        ),
        earthquake_coverage=TargetFields(
            target_keyword="Earthquake Endorsement", first_index=0, second_index=0
        ),
        overland_water=TargetFields(
            target_keyword="Overland Water", first_index=0, second_index=0
        ),
        condo_deductible=TargetFields(
            target_keyword="Condominium Corporation Deductible",
            first_index=0,
            second_index=0,
        ),
        service_line=TargetFields(
            target_keyword="Service Line Coverage", first_index=0, second_index=0
        ),
        premium_amount=TargetFields(
            target_keyword="TOTAL", first_index=0, second_index=remaining_index
        ),
    )._asdict(),
    "Family": FormFields(
        name_and_address=TargetFields(
            target_coordinates=(
                35.49599838256836,
                153.38099670410156,
                150,
                228.67138671875,
            )
        ),
        policy_number=TargetFields(
            target_keyword="POLICY NUMBER",
            target_coordinates=(
                -1.08001708984375,
                11.030815124511719,
                8.86395263671875,
                25.993423461914062,
            ),
        ),
        effective_date=TargetFields(
            target_keyword="EFFECTIVE DATE",
            target_coordinates=(
                -1.00799560546875,
                20.16899871826172,
                24.568084716796875,
                11.449417114257812,
            ),
        ),
        risk_address=TargetFields(
            target_keyword="LOCATION OF INSURED PROPERTY:",
            first_index=0,
            second_index=1,
        ),
        form_type=TargetFields(
            target_keyword="All Perils:", first_index=0, second_index=0
        ),
        risk_type=TargetFields(
            target_keyword="POLICY TYPE",
            target_coordinates=(
                -1.08001708984375,
                11.097042083740234,
                -22.849456787109375,
                0.0,
            ),
        ),
        number_of_families=TargetFields(
            target_keyword="RENTAL SUITES", first_index=0, second_index=0
        ),
        earthquake_coverage=TargetFields(
            target_keyword="EARTHQUAKE PROPERTY LIMITS",
            target_coordinates=(
                9.360000610351562,
                37.858428955078125,
                -144.78810119628906,
                39.08734130859375,
            ),
        ),
        overland_water=TargetFields(
            target_keyword="Overland Water", first_index=0, second_index=0
        ),
        condo_deductible=TargetFields(
            target_keyword="Deductible Coverage:", first_index=0, second_index=0
        ),
        service_line=TargetFields(
            target_keyword="Service Lines", first_index=0, second_index=0
        ),
        premium_amount=TargetFields(
            target_keyword="RETURN THIS PORTION WITH PAYMENT",
            target_coordinates=(
                5.592010498046875,
                -22.78656005859375,
                -116.08038330078125,
                -22.19720458984375,
            ),
        ),
    )._asdict(),
    "Intact": FormFields(
        name_and_address=TargetFields(
            target_coordinates=(
                49.650001525878906,
                152.64999389648438,
                214.1199951171875,
                200.99000549316406,
            )
        ),
        policy_number=TargetFields(
            target_keyword="Policy Number", first_index=1, second_index=0
        ),
        effective_date=TargetFields(
            target_keyword="Policy Number", first_index=1, second_index=1
        ),
        risk_address=TargetFields(
            target_keyword="Property Coverage (",
            target_coordinates=(0.0, 16.902496337890625, 0.0, 0.0),
        ),
        form_type=TargetFields(
            target_keyword="Property Coverage (",
            target_coordinates=(110.64999389648438, 0.0, 0, -10.03228759765625),
        ),
        risk_type=TargetFields(
            target_keyword="Property Coverage (",
            target_coordinates=(110.64999389648438, 0.0, 0, -10.03228759765625),
        ),
        number_of_families=TargetFields(
            target_keyword="Families",
            target_coordinates=(0, 18.699981689453125, 0, 10.54998779296875),
        ),
        premium_amount=TargetFields(
            target_keyword="Total for Policy", first_index=0, second_index=1
        ),
        earthquake_coverage=TargetFields(
            target_keyword="Earthquake Damage Assumption",
            target_coordinates=(
                46.94398498535156,
                11.1737060546875,
                35,
                12.26593017578125,
            ),
        ),
        overland_water=TargetFields(
            target_keyword="Enhanced Water Damage", first_index=0, second_index=0
        ),
        condo_deductible=TargetFields(
            target_keyword="Deductible Coverage:", first_index=0, second_index=0
        ),
        ground_water=TargetFields(
            target_keyword="Ground Water", first_index=0, second_index=0
        ),
        condo_earthquake_deductible=TargetFields(
            target_keyword="Additional Loss Assessment", first_index=0, second_index=0
        ),
        service_line=TargetFields(
            target_keyword="Water and Sewer Lines", first_index=0, second_index=0
        ),
    )._asdict(),
    "Wawanesa": FormFields(
        name_and_address=TargetFields(
            target_coordinates=(36.0, 122.4298095703125, 200, 180)
        ),
        policy_number=TargetFields(
            target_keyword="NAMED INSURED AND ADDRESS", first_index=4, second_index=1
        ),
        effective_date=TargetFields(
            target_keyword="NAMED INSURED AND ADDRESS", first_index=6, second_index=1
        ),
        risk_address=TargetFields(
            target_keyword="Location Description",
            first_index=1,
            second_index=1,
            append_duplicates=True,
        ),
        form_type=TargetFields(
            target_keyword="subject to all conditions of the policy.",
            first_index=3,
            second_index=0,
            append_duplicates=True,
        ),
        risk_type=TargetFields(
            target_keyword="Risk Type",
            first_index=1,
            second_index=2,
            append_duplicates=True,
        ),
        number_of_families=TargetFields(
            target_keyword="Number of Families",
            first_index=0,
            second_index=1,
            append_duplicates=True,
        ),
        number_of_units=TargetFields(
            target_keyword="Number of Units",
            first_index=0,
            second_index=3,
            append_duplicates=True,
        ),
        premium_amount=TargetFields(
            target_keyword="Total Policy Premium", first_index=0, second_index=1
        ),
        earthquake_coverage=TargetFields(
            target_keyword="Earthquake Coverage", first_index=0, second_index=0
        ),
        overland_water=TargetFields(
            target_keyword="Overland Water", first_index=0, second_index=0
        ),
        condo_deductible=TargetFields(
            target_keyword="Condominium Deductible Coverage-",
            target_coordinates=(
                350.89600372314453,
                0.0,
                103.84786987304688,
                -9.600006103515625,
            ),
        ),
        condo_earthquake_deductible=TargetFields(
            target_keyword="Condominium Deductible Coverage Earthquake",
            target_coordinates=(
                350.89600372314453,
                0.0,
                95.4251708984375,
                -9.5999755859375,
            ),
        ),
        tenant_vandalism=TargetFields(
            target_keyword="Vandalism by Tenant Coverage -",
            first_index=0,
            second_index=0,
        ),
        service_line=TargetFields(
            target_keyword="Service Line Coverage", first_index=0, second_index=0
        ),
    )._asdict(),
    "ICBC": FormFields(
        licence_plate=TargetFields(
            target_keyword=re.compile(r"(?<!Previous )\bLicence Plate Number\b"),
            first_index=0,
            second_index=0,
        ),
        transaction_type=TargetFields(
            target_keyword="Transaction Type", first_index=0, second_index=0
        ),
        name_code=TargetFields(
            target_coordinates=(
                198.0,
                761.0403442382812,
                255.010986328125,
                769.977294921875,
            )
        ),
        transaction_timestamp=TargetFields(
            target_keyword="Transaction Timestamp", first_index=0, second_index=0
        ),
        insured_name=TargetFields(
            target_keyword="Name of Insured (surname followed by given name(s))",
            first_index=0,
            second_index=1,
        ),
        owner_name=TargetFields(target_keyword="Owner ", first_index=0, second_index=1),
        applicant_name=TargetFields(
            target_keyword="Applicant", first_index=0, second_index=1
        ),
        validation_stamp=TargetFields(
            target_keyword="NOT VALID UNLESS STAMPED BY",
            target_coordinates=(
                -4.247998046875011,
                13.768768310546875,
                1.5807250976562273,
                58.947509765625,
            ),
        ),
        customer_copy=TargetFields(
            target_keyword="Customer Copy",
            target_coordinates=(
                36.0,
                761.039794921875,
                560.1806640625,
                769.977294921875,
            ),
        ),
        time_of_validation=TargetFields(
            target_keyword="TIME OF VALIDATION",
            target_coordinates=(0.0, 10.354278564453125, 0.0, 40),
        ),
    )._asdict(),
}


def ff(dictionary):
    for key, value in dictionary.items():
        if value and isinstance(value[0], str):
            dictionary[key] = [value]
    return dictionary


def find_index(regex, dict_item):
    if dict_item and isinstance(dict_item, list):
        for index, string in enumerate(dict_item):
            if re.search(regex, string):
                return index
        return -1


def join_and_format_names(names):
    if len(names) == 1:
        return names[0].title()
    else:
        return (
            ", ".join(names[:-1]).strip().title() + " and " + names[-1].strip().title()
        )


def address_one_title_case(sentence):
    return " ".join(
        word.lower() if word[-2:] in {"th", "rd", "nd", "st"} else word.capitalize()
        for word in sentence.split()
    )


def address_two_title_case(strings_list):
    words = strings_list.split()
    capitalized_words = [
        word.strip().capitalize() if len(word) > 2 else word for word in words
    ]
    return " ".join(capitalized_words)


def risk_address_title_case(address):
    parts = address.split()
    last_part = parts[-1]
    if len(last_part) == 2:
        last_part = last_part.upper()
    titlecased_address_parts = []
    for part in parts:
        if part[:-2].isdigit() and part[-2:].lower() in ["th", "rd", "nd", "st"]:
            titlecased_address_parts.append(part.lower())
        else:
            titlecased_address_parts.append(part.title())
    titlecased_address = " ".join(titlecased_address_parts[:-1]) + " " + last_part
    return titlecased_address


def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


def find_matching_paths(target_filename, paths):
    matching_paths = [path for path in paths if path.stem.split()[0] == target_filename]
    return matching_paths
