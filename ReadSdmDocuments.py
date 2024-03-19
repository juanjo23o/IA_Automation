from utilities import extract_infortmation, process_excel
from openpyxl import load_workbook
import pandas as pd
import re
import os
import re
import shutil

carpeta_excel = 'SupportDocuments'
archivos_excel = os.listdir(carpeta_excel)

accounts_checked = 0
accounts_unchecked = 0

account_details = []


for i in archivos_excel:

        account = fr'SupportDocuments\{i}'

        try:
            wb = load_workbook(account)
        except:continue

        # sheet_1 = '3-15-2024'
        # sheet_2 = '03-15-2024'
        # sheet_name = ''
        # flag_account = True
        # # Obtén los nombres de las hojas
        # sheet_names = wb.sheetnames

        # # Imprime los nombres de las hojas
        # for name in sheet_names:
        #     if sheet_1 == name.strip() or sheet_2 == name.strip():
        #         sheet_name = name
        #         sheet = wb[sheet_name]
        #         break
        #     else:
        #         shutil.move(account, r"wrong_cycle")
        #         flag_account = False
        #         break
        # if flag_account == True:
        #     pass
        # else:
        #     continue

        try:
            try:
                sheet = wb['3-15-2024']
                sheet_name = '3-15-2024'
            except:
                try:
                    sheet = wb['3-15-2024 ']
                    sheet_name = '3-15-2024 '
                except:
                    try:
                        sheet = wb['3-15-2024  ']
                        sheet_name = '3-15-2024  '
                    except:
                        sheet = wb['03-15-2024']
                        sheet_name = '03-15-2024'
        except:
            shutil.move(account, r"wrong_cycle")
            continue
        
        
        try:
        # archivo_excel = account
            df = pd.read_excel(account, sheet_name=sheet_name)
            archivo_txt = 'example.txt'
            df.to_csv(archivo_txt, sep='\t', index=False)

            with open('example.txt', 'r', encoding='utf-8') as archivo:
                contenido = archivo.read()
                archivo.close()

            company_name = re.findall(r'Company:\s*(.*)', contenido)[0].strip()
            cleaned_company_name = re.split(r'\t', company_name)[0]

            results = process_excel(sheet)

            df = pd.DataFrame(results[0], columns=results[1])

            flag = 'SDM'

            employees = extract_infortmation(df, 'SDM', cleaned_company_name, account, results[2]['rate'], results[2]['equipment'], results[2]['active_date'], results[2]['rate_adjustment'], results[2]['credit_days'], account.split('\\')[1], results[2]['set_up_fee'], results[2]['bonus'])

            account_details.append(employees)
            with open('accounts_processed_sdm.txt', 'a', encoding='utf-8') as archivo:
                archivo.write(account.split('\\')[1].split("Support")[0] + '\n')
            accounts_checked += 1
        
        except:
            print(f'esta cuenta no pudo ser procesada: {account}')
            shutil.move(account, r"accounts_to_review")
            accounts_unchecked += 1
    
    



df_list = []
for group in account_details:
    company_df = pd.DataFrame()
    for employee in group:
        employee_data = list(employee.values())[0]
        employee_df = pd.json_normalize(employee_data)
        company_df = pd.concat([company_df, employee_df], axis=0, ignore_index=True)
    
    df_list.append(company_df)

final_df = pd.concat(df_list, axis=0, ignore_index=True)

final_df.to_excel('output.xlsx', index=False)

print(accounts_checked)
print(accounts_unchecked)