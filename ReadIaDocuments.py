from utilities import extract_infortmation, process_excel
from openpyxl import load_workbook
import pandas as pd
import re
import os
import re
import shutil


carpeta_excel = 'DocumentsGeneratedIA'
archivos_excel = os.listdir(carpeta_excel)

modified_columns = list()
accounts_checked = 0
accounts_unchecked = 0
accounts_processed = 'accounts_processed.txt'
account_details = []


for i in archivos_excel:
    try:

        account = fr'DocumentsGeneratedIA\{i}'

        # archivo_excel = account
        df = pd.read_excel(account)
        archivo_txt = 'example2.txt'
        df.to_csv(archivo_txt, sep='\t', index=False)

        with open('example2.txt', 'r', encoding='utf-8') as archivo:
            contenido = archivo.read()
            archivo.close()


        company_name = re.findall(r'Company:\s*(.*)', contenido)[0].strip()
        cleaned_company_name = re.split(r'\t', company_name)[0]

        wb = load_workbook(account)
        sheet = wb['Sheet1']
        
        results = process_excel(sheet)

        df = pd.DataFrame(results[0], columns=results[1])

        flag = 'IA'

        employees = extract_infortmation(df, 'SDM', cleaned_company_name, account, results[2]['rate'], results[2]['equipment'], results[2]['active_date'], results[2]['rate_adjustment'], results[2]['credit_days'], account.split('\\')[1], results[2]['set_up_fee'], results[2]['bonus'])

        if employees == False:
            accounts_unchecked += 1
            continue
        else:
            account_details.append(employees)
            with open('accounts_processed.txt', 'a', encoding='utf-8') as archivo:
                account_name = account.split('\\')[1]
                archivo.write(account_name.split('_')[0] + '\n')
            accounts_checked += 1
    
    except:
        shutil.move(account, r"C:\Users\Usuario\Documents\invoicing-automation\accounts_to_review")
        print(f"-- The following account can't be processed {cleaned_company_name} --")
        accounts_unchecked += 1


df_list = []
for group in account_details:
    company_df = pd.DataFrame()
    for employee in group:
        employee_data = list(employee.values())[0]
        employee_df = pd.json_normalize(employee_data)
        company_df = pd.concat([company_df, employee_df], axis=0, ignore_index=True)
    
    df_list.append(company_df)

# Consolidar todos los DataFrames en uno solo
final_df = pd.concat(df_list, axis=0, ignore_index=True)

# Guardar en un archivo Excel
final_df.to_excel('output2.xlsx', index=False)

print(accounts_checked)
print(accounts_unchecked)