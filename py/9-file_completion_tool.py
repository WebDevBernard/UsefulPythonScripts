import pandas as pd
from coordinates import base_dir
from sort_renewal_list import sort_renewal_list
from file_writing_functions import renewal_letter, renewal_letter_manual

input_dir = base_dir / "input"
output_dir = base_dir / "output"
output_dir.mkdir(exist_ok=True)
pdf_files = input_dir.glob("*.pdf")
excel_path = base_dir / "input.xlsx"  # name of Excel

def get_excel_data(excel_path):
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

excel_data = get_excel_data(excel_path)

if __name__ == "__main__":
    df_excel = pd.read_excel(excel_path, sheet_name=0, header=None)
    task = df_excel.at[6, 1]
    if task == "Auto Generate Renewal Letter From Policy":
        renewal_letter(excel_path)
    elif task == "Manually Generate Renewal Letter from Data Entry":
        renewal_letter_manual(excel_data)
    elif task == "Sort Renewal List from Excel File":
        sort_renewal_list()
