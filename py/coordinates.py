import re

postal_code_regex = r"[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d$"
dollar_regex = r"\$([\d,]+)"

doc_type = {
    "Aviva": ["Company", (171.36000061035156, 744.800048828125, 204.39999389648438, 752.7999877929688)],
    "Family": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
    "Intact": ["Policy", (45.099998474121094, 23.674999237060547, 69.10299682617188, 36.04100036621094)],
    "Wawanesa": ["Policy", (36.0, 187.6835479736328, 62.479373931884766, 197.7382354736328)],
    "Invoice": ["CUSTOMER STATEMENT", (179.52000427246094, 96.41597747802734, 396.89202880859375, 116.48722839355469)]
}

keyword = {
    "Aviva": ["C e r t i f i c a t e  O f  P r o p e r t y  I n s u r a n c e", (90.23999786376953, 99.76000213623047, 509.0579833984375, 109.760009765625)],
    "Intact": ["BROKER COPY", (37.650001525878906, 236.69998168945312, 197.68801879882812, 261.4319763183594)],
    "Wawanesa": [re.compile(r"\w{3} \d{2}, \d{2}"), (36.0, 762.829833984375, 576.001220703125, 778.6453857421875)],
}

target_dict = {
    "Aviva": {
        "name_and_address": (80.4000015258789, 202.239990234375, 250, 280),
        "policy_number": [["Policy Number", (267.11999893188477, 10.15997314453125, -202.8189697265625, 9.16009521484375)]],
        "effective_date": ["Policy Effective From:", 0, 1],
        "policy_deductible": ["LOCATION 1", 0, 2],
        "risk_address": ["FORM", -1, slice(*map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(':')))],
        "aviva_form_type": ["FORM", 0, 0, True],
        "aviva_number_of_families": ["Family", 0, 1, True],
        "earthquake_coverage": ["Earthquake Endorsement", 0, 0],
        "overland_water": ["Overland Water", 0, 0],
        "condo_deductible_coverage": ["Condominium Corporation Deductible", 0, 0],
        "premium_amount": ["TOTAL", 0, slice(*map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(':')))],
    },
    "Family": {
        "name_and_address": (37.72800064086914, 170.62953186035156, 122.84105682373047, 220.67938232421875),
        "policy_number": [["POLICY NUMBER", (0, 10.913284301757812, 0, 12.68017578125)]],
        "effective_date": [["EFFECTIVE DATE", (0, 11.345291137695312, 0, 13.1121826171875)]],
        "risk_address": ["LOCATION OF INSURED PROPERTY:", 0, 1],
        "form_type": ["All Perils:", 0, 0],
        "risk_type": [["POLICY TYPE", (0, 11.633319854736328, 15, 13.40020751953125)]],
        "policy_deductible": ["All Property Deductible.", 0, 0],
        "number_of_families": ["RENTAL SUITES", 0, 0],
        "premium_amount": [["RETURN THIS PORTION WITH PAYMENT", (0, -19.8719482421875, 0, -19.8719482421875)]],
        "condo_deductible_coverage": ["Deductible Coverage:", 0, 0],
    },
    "Intact": {
        "name_and_address": (49.650001525878906, 152.64999389648438, 214.1199951171875, 200.99000549316406),
        "policy_number": ["Policy Number", 1, 0],
        "effective_date": ["Policy Number", 1, 1],
        "risk_address": ["Property Coverage", 0, 1],
        "form_type": ["Property Coverage", 0, 0],
        "policy_deductible": ["All Property Deductible.", 0, 0],
        "number_of_families": [["Families", (0, 18.699981689453125, 0, 10.54998779296875)]],
        "premium_amount": ["Total for Policy", 0, 1],
        "condo_earthquake_deductible": ["Additional Loss Assessment Extension", 0, 0],
    },
   "Wawanesa": {
        "name_and_address": (36.0, 122.4298095703125, 137.55322265625, 169.76731872558594),
        "policy_number": ["NAMED INSURED AND ADDRESS", 4, 1],
        "effective_date": ["NAMED INSURED AND ADDRESS", 6, 1],
        "risk_address": ["Location Description", 1, 1],
        "form_type": ["subject to all conditions of the policy.", 3, 0],
        "risk_type": ["Risk Type", 1, 2],
        "policy_deductible": ["Original Deductible", 0, 1],
        "number_families": ["Number of Families", 0, 1],
        "premium_amount": ["Total Policy Premium", 0, 1],
        "condo_deductible_coverage": ["Condominium Deductible Coverage", 1, 0],
        "condo_earthquake_deductible": ["Condominium Deductible Coverage Earthquake", 1, 0],
    },
    "Invoice": {
    "name_and_address": (46.560001373291016, 142.80056762695312, 179.52000427246094, 180.94363403320312)
    }
}
dict_of_keywords = {
    "One": 1,
    "Two": 2,
    "Three": 3,
    "COMPREHENSIVE FORM": "Comprehensive",
    "HOMEOWNERS": "Homeowners",
    "CONDOMINIUM": "Condominium",
    "RENTED DWELLING" : "Rented Dwelling",
    "FIRE & EXTENDED COVERAGE FORM": "Fire and EC"
}

filename = {
    "RENEWAL LETTER": "Renewal Letter - Copy.docx",
}