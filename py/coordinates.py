import re
# the pdf type

doc_type = {
    "Aviva": ["Company", (171.36000061035156, 744.800048828125, 204.39999389648438, 752.7999877929688)],
    "Family": ["Agent", (26.856000900268555, 32.67083740234375, 48.24102783203125, 40.33245849609375)],
    "Intact": ["Policy", (45.099998474121094, 23.674999237060547, 69.10299682617188, 36.04100036621094)],
    "Wawanesa": ["Policy", (36.0, 187.6835479736328, 62.479373931884766, 197.7382354736328)],
    "Invoice": ["CUSTOMER STATEMENT", (179.52000427246094, 96.41597747802734, 396.89202880859375, 116.48722839355469)]
}

# Keyword to find all the pages with broker copy or coverage summary
# set searches not used, used if you are looking for a page to stop on starting from page 1
# string search returns a single page, used for aviva
# list searches is for intact
# regex searches is for wawanesa

keyword = {
    "Aviva": ["C e r t i f i c a t e  O f  P r o p e r t y  I n s u r a n c e", (90.23999786376953, 99.76000213623047, 509.0579833984375, 109.760009765625)],
    "Intact": [["BROKER COPY", "Coverage Summary"], [(265.20001220703125, 764.2749633789062, 347.6670227050781, 773.8930053710938), (37.650001525878906, 236.69998168945312, 197.68801879882812, 261.4319763183594)]],
    "Wawanesa": [re.compile(r"\w{3} \d{2}, \d{2}"), (36.0, 762.829833984375, 576.001220703125, 778.6453857421875)],
}

# number to index in the outer list and number to index in the inner list after finding matching keywords
# tuple only searches page 1
# nest list searches using crop relative to input_coords
# string searches a match from dictionary
target_dict = {
    "Aviva": {
        "name_and_address": (80.4000015258789, 202.239990234375, 193.14999389648438, 239.60003662109375),
        "policy_number": ["POLICY NUMBER:", 0, 1],
        "effective_date": ["Policy Effective From:", 0, 1],
        "risk_address": ["Location 1", 0, slice(*map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(':')))],
        "risk_address_2": ["Location 2", 0, slice(*map(lambda x: int(x.strip()) if x.strip() else None, "1:".split(':')))],
        "risk_location_1_premium": ["TOTAL", 0, 1],
        "risk_location_2_premium": ["TOTAL", 0, 2],
        "number_families": ["Family", 0, 1],
        "form_type": ["FORM", 0, 0],
        "deductible": ["LOCATION 1", 0, 2],
        "coverage_a": ["Coverage A", 0, 1],
        "legal_liability": ["Coverage E - Legal Liability", 0, 1],
        "has_earthquake": ["Earthquake Endorsement", 0, 0],
        "overland_water": ["Overland Water", 0, 0],
        "has_sewer_back": ["Overland Water", 0, 0],
    },
    "Family": {
        "name_and_address": (37.72800064086914, 170.62953186035156, 122.84105682373047, 220.67938232421875),
        "policy_number": [["POLICY NUMBER", (0, 10.913284301757812, 0, 12.68017578125)]],
        "effective_date": [["EFFECTIVE DATE", (0, 11.345291137695312, 0, 13.1121826171875)]],
        "premium_amount": [["RETURN THIS PORTION WITH PAYMENT", (0, -19.8719482421875, 0, -19.8719482421875)]],
        "risk_address": ["LOCATION OF INSURED PROPERTY:", 0, 1],
        "form_type": ["PREMIUM MUST BE RECEIVED BY THE EFFECTIVE DATE TO AVOID CANCELLATION OF COVERAGE.", 13, 0],
        "deductible": ["All Property Deductible.", 0, 0],
        "all_perils": ["All Perils:", 0, 0],
        "specified_perils": ["Specified Perils:", 0, 0],
        "has_earthquake": ["COVERAGES AND LIMITS OF INSURANCE - EARTHQUAKE", 1, 0],
        "earthquake_building_only": ["COVERAGES AND LIMITS OF INSURANCE - EARTHQUAKE", 3, 0],
        "legal_liability": ["Coverage E - Legal Liability", 0, 1],
        "number_of_families": ["RENTAL SUITES", 0, 0],
        "coverage_d": ["Personal Property: Unlimited", 1, 0],
        "claim_forgiveness": ["Claim Forgiveness", 0, 0],
    },
    "Intact": {
        "name_and_address": (49.650001525878906, 152.64999389648438, 214.1199951171875, 200.99000549316406),
        "policy_number": ["Policy Number", 1, 0],
        "effective_date": ["Policy Number", 1, 1],
        "risk_address": ["Property Coverage", 0, 1],
        "form_type": ["Property Coverage", 0, 0],
        "risk_type": ["Property Coverage", 0, 0],
        "number_of_families": [["Families", (0, 18.699981689453125, 0, 10.54998779296875)]],
        "has_overland_water": ["Overland Water", 0, 0],
        "has_sewer_back": ["Overland Water", 0, 0],
        "has_ground_water": ["Overland Water", 0, 0],
        "has_earthquake": ["", 1, 1],
        "deductible": ["All Property Deductible.", 0, 0],
        "legal_liability": ["Coverage E - Legal Liability", 0, 1],
        "coverage_a": ["Dwelling Building", 0, 3],
    },
   "Wawanesa": {
        "name_and_address": (36.0, 122.4298095703125, 137.55322265625, 169.76731872558594),
        "policy_number": ["NAMED INSURED AND ADDRESS", 4, 1],
        "effective_date": ["NAMED INSURED AND ADDRESS", 6, 1],
        "number_of_locations": ["Location Description", 1, 0],
        "risk_address": ["Location Description", 1, 1],
        "premium_amount": ["Total Policy Premium", 0, 1],
        "number_families": ["Number of Families", 0, 1],
        "form_type": ["subject to all conditions of the policy.", 3, 0],
        "risk_type": ["Risk Type", 1, 2],
        "deductible": ["Original Deductible", 0, 1],
        "coverage_a": ["Coverage A - Dwelling", 0, 1],
        "legal_liability": ["Coverage E - Legal Liability", 0, 1],
        "bylaw_coverage": ["Building Bylaws Coverage", 0, 1],
        "service_line": ["Service Line Coverage", 0, 1],
        "condo_deductible_coverage": ["Condominium Deductible Coverage", 1, 0],
        "condo_earthquake_deductible": ["Condominium Deductible Coverage Earthquake", 1, 0],
        "has_earthquake": ["Condominium Unit Owners Earthquake Coverage", 1, 1],
        "has_overland_water": ["Overland Water", 0, 0],
        "has_tenant_vandalism": ["Vandalism by Tenant", 0, 1],
    },
    "Invoice": {
    "name_and_address": (46.560001373291016, 142.80056762695312, 179.52000427246094, 180.94363403320312)
    }
}
