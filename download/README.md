# UsefulPythonScripts

- In a new folder, create another folder inside named `input`
- Drag and drop all the templates `.docx` and `.pdf` into `input`
- In the Root folder create or drag and drop the `input.xlsx` file
- Add the following columns `type` `insured_name` `additional_insured` `policy_number` `effective_date` `mailing_address` `risk_address` `insurer` `broker_name`
- In program folder, create short cut of `program.exe` and drag and drop into the root folder
- Run `program.exe` shortcut
- The `docx` or ` pdf`` is generated in the  `output` folder

## How to use:

- You do not need to fill out risk_address, if mailing address is the same. The program automatically assumes it's a home, tenant or condo policy.
- Filling out risk_address along with mailing address will auto generate a rental questionnaire depending on the `insurer`
- Under `insurer`, typing `gore mutual` `intact` `optimum west` `wawanesa` will generate the corresponding rental questionnaire
- To generate a `wawanesa` rented dwelling questionnaire (house, not condo), under `type`, type in `revenue`
- To generate a binder and binder invoice, under `type` type in `binder`
- To generate a lapse or cancel letter, under type `type` type in `cancel` or type in `lapse`
- The following columns: `type` `insured_name` `additional_insured` `insurer` `broker_name` will auto captialize the first letter of each word. There is no need to capitalize any names.
- Under `broker_name` enter the CSR responsible for this client. You can also leave this blank and it will default to the drop down box where you can manually select the CSR once the word docx is generated
- Important - double check everything before attaching: coverage decline form uses fixed checkboxes based on commonly decline coverages, questionnaires will only fill out names and effective dates
- Troubleshooting - if the program does not work, it is likely you have a word or pdf document open. The program is trying to overwrite the word or pdf file, and it cannot work if the file is open. Close the word or pdf file and it should work.
