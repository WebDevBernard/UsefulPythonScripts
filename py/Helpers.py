import re
from collections import namedtuple

postal_code_regex = re.compile(r"([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d)$")
dollar_regex = re.compile(r"\$([\d,]+)")
date_regex = re.compile(r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}|(Jan(uary)?|Feb(ruary)?"
                        r"|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?"
                        r"|Dec(ember)?)\s+\d{1,2},\s+\d{4}")
address_regex = re.compile(r"(?!.*\bltd\.)((po box)|(unit)|\d+\s+)", flags=re.IGNORECASE)
and_regex = re.compile(r'&|\b(and)\b', flags=re.IGNORECASE)

form_fields = ["insurer", "name_and_address", "policy_number", "effective_date", "risk_address", "form_type",
                  "risk_type", "number_of_families", "number_of_units", "premium_amount", "condo_deductible",
                  "condo_earthquake_deductible", "earthquake_coverage", "overland_water", "ground_water", "tenant_vandalism",
                  "service_line"]
FormFields = namedtuple("FormFields", form_fields, defaults=(None,) * len(form_fields))
target_fields = ["target_keyword", "first_index", "second_index", "target_coordinates", "append_duplicates", "join_list"]
TargetFields = namedtuple("TargetFields", target_fields, defaults=(None,) * len(target_fields))
remaining_index = slice(*map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(':')))
target_dict = {
    "Aviva":
        FormFields(
            # "name_and_address" is the target_key, first_index the outer nested list, inner_index is the inner list
            name_and_address=TargetFields(target_coordinates=(80.4000015258789, 202.239990234375, 250, 280)),
            policy_number=TargetFields(target_keyword="Policy Number", target_coordinates=(
                267.11999893188477, 10.15997314453125, -202.8189697265625, 9.16009521484375)),
            effective_date=TargetFields(target_keyword="Policy Effective From:", first_index=0, second_index=1),
            risk_address=TargetFields(target_keyword="FORM", first_index=-1, second_index=remaining_index, join_list=True),
            form_type=TargetFields(target_keyword=re.compile(r".* - .*form$", flags=re.IGNORECASE), first_index=0,
                                   second_index=0,
                                   append_duplicates=True),
            risk_type=TargetFields(target_keyword=re.compile(r".* - .*form$", flags=re.IGNORECASE), first_index=0,
                                   second_index=0,
                                   append_duplicates=True),
            number_of_families=TargetFields(target_keyword="Extended Liability", first_index=1, second_index=0),
            earthquake_coverage=TargetFields(target_keyword="Earthquake Endorsement", first_index=0, second_index=0),
            overland_water=TargetFields(target_keyword="Overland Water", first_index=0, second_index=0),
            condo_deductible=TargetFields(target_keyword="Condominium Corporation Deductible", first_index=0,
                                          second_index=0),
            service_line=TargetFields(target_keyword="Service Line Coverage", first_index=0, second_index=0),
            premium_amount=TargetFields(target_keyword="TOTAL", first_index=0, second_index=remaining_index)
        )._asdict(),
    "Family":
        FormFields(
            # "name_and_address" is the target_key
            name_and_address=TargetFields(
                target_coordinates=(37.72800064086914, 170.62953186035156, 150, 220.67938232421875)),
            policy_number=TargetFields(target_keyword="POLICY NUMBER",
                                       target_coordinates=(0, 10.913284301757812, 0, 12.68017578125)),
            effective_date=TargetFields(target_keyword="EFFECTIVE DATE",   target_coordinates=(0, 11.345291137695312, 0, 13.1121826171875)),
            risk_address=TargetFields(target_keyword="LOCATION OF INSURED PROPERTY:", first_index=0, second_index=1),
            form_type=TargetFields(target_keyword="All Perils:", first_index=0, second_index=0),
            risk_type=TargetFields(target_keyword="POLICY TYPE",
                                   target_coordinates=(0, 11.633319854736328, 15, 13.40020751953125)),
            number_of_families=TargetFields(target_keyword="RENTAL SUITES", first_index=0, second_index=0),
            earthquake_coverage=TargetFields(target_keyword="EARTHQUAKE PROPERTY LIMITS", target_coordinates=(
            46.94398498535156, 11.1737060546875, 35, 12.26593017578125)),
            overland_water=TargetFields(target_keyword="Overland Water", first_index=0, second_index=0),
            condo_deductible=TargetFields(target_keyword="Deductible Coverage:", first_index=0,
                                          second_index=0),
            service_line=TargetFields(target_keyword="Service Lines", first_index=0, second_index=0),
            premium_amount=TargetFields(target_keyword="RETURN THIS PORTION WITH PAYMENT",
                                        target_coordinates=(0, -19.8719482421875, 0, -19.8719482421875))
        )._asdict(),
    "Intact":
        FormFields(
            # "name_and_address" is the target_key
            name_and_address=TargetFields(
                target_coordinates=(49.650001525878906, 152.64999389648438, 214.1199951171875, 200.99000549316406)),
            policy_number=TargetFields(target_keyword="Policy Number",
                                       first_index=1, second_index=0),
            effective_date=TargetFields(target_keyword="Policy Number", first_index=1, second_index=1),
            risk_address=TargetFields(target_keyword="Property Coverage (", first_index=0, second_index=remaining_index, append_duplicates=True),
            form_type=TargetFields(target_keyword="Property Coverage (", first_index=0, second_index=0),
            risk_type=TargetFields(target_keyword="Property Coverage (", first_index=0, second_index=0),
            number_of_families=TargetFields(target_keyword="Families", target_coordinates=(0, 18.699981689453125, 0, 10.54998779296875)),
            premium_amount=TargetFields(target_keyword="Total for Policy",
                                        first_index=0, second_index=1),
            earthquake_coverage=TargetFields(target_keyword="Earthquake Damage Assumption", target_coordinates=(
                46.94398498535156, 11.1737060546875, 35, 12.26593017578125)),
            overland_water=TargetFields(target_keyword="Enhanced Water Damage", first_index=0, second_index=0),
            condo_deductible=TargetFields(target_keyword="Deductible Coverage:", first_index=0,
                                          second_index=0),
            ground_water=TargetFields(target_keyword="Ground Water", first_index=0, second_index=0),
            condo_earthquake_deductible=TargetFields(target_keyword="Additional Loss Assessment", first_index=0, second_index=0),
            service_line=TargetFields(target_keyword="Water and Sewer Lines", first_index=0, second_index=0),

        )._asdict(),
    "Wawanesa":
        FormFields(
            # "name_and_address" is the target_key
            name_and_address=TargetFields(
                target_coordinates=(36.0, 122.4298095703125, 200, 180)),
            policy_number=TargetFields(target_keyword="NAMED INSURED AND ADDRESS",
                                       first_index=4, second_index=1),
            effective_date=TargetFields(target_keyword="NAMED INSURED AND ADDRESS", first_index=6, second_index=1),
            risk_address=TargetFields(target_keyword="Location Description", first_index=1, second_index=1,
                                      append_duplicates=True),
            form_type=TargetFields(target_keyword="subject to all conditions of the policy.", first_index=3, second_index=0, append_duplicates=True),
            risk_type=TargetFields(target_keyword="Risk Type",
                                   first_index=1, second_index=2, append_duplicates=True),
            number_of_families=TargetFields(target_keyword="Number of Families",
                                            first_index=0, second_index=1, append_duplicates=True),
            number_of_units=TargetFields(target_keyword="Number of Units",
                                            first_index=0, second_index=3, append_duplicates=True),
            premium_amount=TargetFields(target_keyword="Total Policy Premium",
                                        first_index=0, second_index=1),
            earthquake_coverage=TargetFields(target_keyword="Earthquake Coverage", first_index=0, second_index=0),
            overland_water=TargetFields(target_keyword="Overland Water", first_index=0, second_index=0),
            condo_deductible=TargetFields(target_keyword="Condominium Deductible Coverage-", first_index=0,
                                          second_index=1),
            condo_earthquake_deductible=TargetFields(target_keyword="Condominium Deductible Coverage Earthquake", first_index=1,
                                                     second_index=0),
            tenant_vandalism=TargetFields(target_keyword="Vandalism by Tenant Coverage -", first_index=0, second_index=0),
            service_line=TargetFields(target_keyword="Service Line Coverage", first_index=0, second_index=0),

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
        return ', '.join(names[:-1]).strip().title() + ' and ' + names[-1].strip().title()


def address_one_title_case(sentence):
    return ' '.join(
        word if word[-2:] in {"th", "rd", "nd", "st"} else word.capitalize()
        for word in sentence.split()
    )

def address_two_title_case(strings_list):
    words = strings_list.split()
    capitalized_words = [word.strip().capitalize() if len(word) > 2 else word for word in words]
    return ' '.join(capitalized_words)