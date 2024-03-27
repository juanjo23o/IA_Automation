
from datetime import datetime
import pandas as pd
import pycountry
import re


def process_excel(sheet):
    modified_columns = {
        'id':'',
        'rate':'',
        'set_up_fee':'',
        'active_date':'',
        'bonus':'',
        'equipment':'',
        'ot_hours':'',
        'ot_amount':'',
        'rate_adjustment':'',
        'credit_days':''
    }
    found_workplace = False
    data = []
    columns = []
    flag_column = 0
    for row in sheet.iter_rows(values_only=True):
        try:
            row_flag = row[0].strip()
        except:
            row_flag = row[0]
        if found_workplace:
            try:
                if ' ' in row_flag or row_flag == '':
                    found_workplace = False
                    continue
                elif row_flag == None:
                    found_workplace = False
                    continue
                elif len(row_flag) > 2:
                    data.append(row)
            except TypeError:
                try:
                    if row[1].strip() != '':
                            data.append(row)
                except:
                    found_workplace = False
                    continue
        elif row_flag == 'Workplace':
            new_row = []
            for i in row:
                try:
                    new_row.append(i.strip().lower())
                except:
                    new_row.append(i)
            rate_adj_date_column = [] 
            for i in new_row:
                try:
                    if 'rate adjustment' in i:
                        rate_adj_date_column.append(i)
                except:
                    rate_adj_date_column.append(i)
            if rate_adj_date_column.count(rate_adj_date_column[0]) >= 1:
                if flag_column >= 1:
                    found_workplace = True
                    data[0] = tuple(row)
                    continue
                data.insert(0, tuple(row))
                found_workplace = True
                flag_column += 1
    columns = data[0]
    new = []
    for i in columns:
        try:
            column = i.strip().lower()
            if 'ID' in i.strip():
                get_column_id = re.findall(r'(id)', column)[0]
                column = get_column_id
                modified_columns['id'] = column
            elif 'monthly' in column or 'amount b' in column:
                try:
                    get_column_rate = re.search(r'(monthly\s+\w+)|(^amount billed\b)', column)[0]
                    if get_column_rate == '':
                        pass
                    else:
                        column = get_column_rate
                        modified_columns['rate'] = column
                except:
                    column = None
            elif 'active' in column:
                get_column_active_date = re.findall(r'(active\s+\w+)', column)[0]
                column = get_column_active_date
                modified_columns['active_date'] = column
            elif 'equipment' in column:
                column_line = "".join(column.splitlines())
                get_column_equipment = re.findall(r'(equipment\s+\w.+)', column)[0].strip()
                column = get_column_equipment
                modified_columns['equipment'] = column
            elif 'rate adjustment' in column:
                get_column_adj_date = re.findall(r'(rate adjustment\s+\w.+)', column)[0]
                column = get_column_adj_date
                modified_columns['rate_adjustment'] = column
            elif 'credit days' in column.replace('  ', ' '):
                get_column_credit_days = re.findall(r'(prorate.+)', column)[0]
                column = get_column_credit_days
                modified_columns['credit_days'] = column
            elif 'set' in column:
                get_column_set_up_fees = re.findall(r'(set.+)', column)[0]
                column = get_column_set_up_fees.replace('  ', ' ')
                modified_columns['set_up_fee'] = column
            elif 'bonus' in column:
                get_column_bonus = re.findall(r'(bonus*.+)',column.replace(' ',''))[0]
                column = get_column_bonus
                modified_columns['bonus'] = column
            elif 'ot hours' in column:
                get_column_hours = re.findall(r'(ot hours)',column)[0]
                column = get_column_hours
                modified_columns['ot_hours'] = column
            elif 'ot amount' in column:
                get_column_amount = re.findall(r'(ot amount)',column)[0]
                column = get_column_amount
                modified_columns['ot_amount'] = column 
            else:...
            new.append(column)
        except:
            new.append(None)
    data.pop(0)
    return data, new, modified_columns



def extract_infortmation(df, flag, *args):
    employees = []
    index = 0
    if flag == 'IA':
        file_name = args[7].split('_')[0]
    elif flag == 'SDM':
        file_name = args[7].split("Support")[0] 
    for i in range(0, len(df)):
        employee = {}
        employee_data = {
            "company_name": '',
            "file_name": '',
            "id": '',
            "name": '',
            "title": '',
            "status": '',
            "rate": '',
            "credit_days": '',
            "prorate_amount": '',
            "set_up_fee": '',
            "bonus_commission": '',
            "ot_hours": '',
            "ot_amount": '',
            "equipment_rental_fee": '',
            "active_date": '',
            "rate_adj_date": '',
            "assigned_date": ''
        }

        
        arguments = [args[0], file_name, str(df.at[i, 'id']), df.at[i, 'name'].lower(), df.at[i, 'title'].lower(), 
                            df.at[i, 'status'].strip().lower(), df.at[i, args[2]], df.at[i, args[6]],
                            df.at[i, "prorate amount"], df.at[i, args[8]], df.at[i, args[9]],
                            df.at[i, args[10]], df.at[i, args[11]], df.at[i, args[3]], df.at[i, args[4]], 
                            df.at[i, args[5].lower()], df.at[i, 'assigned date']]

        for i, key in enumerate(employee_data.keys()):
                try:
                    if pycountry.countries.lookup(arguments[i]):
                        employee_data[key] = 'This column does not exist in the sdm'
                    # else:
                    #     employee_data[key] = arguments[i]
                except:
                    employee_data[key] = arguments[i]

        employee['employee_'+str(index)] = employee_data
        
        if pd.isna(employee['employee_'+str(index)]["credit_days"]):
            employee['employee_'+str(index)]["credit_days"] = ''
        else:
            try:
                employee['employee_'+str(index)]["credit_days"] = abs(int(employee['employee_'+str(index)]["credit_days"]))
            except:
                employee['employee_'+str(index)]["credit_days"] = ''

        
        if pd.isna(employee['employee_'+str(index)]["set_up_fee"]):
            employee['employee_'+str(index)]["set_up_fee"] = ''
        
        if pd.isna(employee['employee_'+str(index)]["bonus_commission"]):
            employee['employee_'+str(index)]["bonus_commission"] = ''

        if pd.isna(employee['employee_'+str(index)]["ot_hours"]):
            employee['employee_'+str(index)]["ot_hours"] = ''
        else:
            try:
                if employee['employee_'+str(index)]["ot_hours"].startswith('='):
                    employee['employee_'+str(index)]["ot_hours"] = int(eval(employee['employee_'+str(index)]["ot_hours"][1:]))
            except:
                employee['employee_'+str(index)]["ot_hours"] = employee['employee_'+str(index)]["ot_hours"]

        try:
            if employee['employee_'+str(index)]["ot_amount"].startswith('='):
                expression = employee['employee_'+str(index)]["ot_amount"][1:]
                workable_days = expression.split('*')[1]
                employee['employee_'+str(index)]["ot_amount"] = round(eval(f"{employee['employee_'+str(index)]['ot_hours']}*{workable_days}"), 2)
        except:
            try:
                employee['employee_'+str(index)]["ot_amount"] = round(employee['employee_'+str(index)]["ot_amount"], 2)
            except:
                employee['employee_'+str(index)]["ot_amount"] = ''

        id = employee['employee_'+str(index)]["id"]
        try:
            id_ = re.search(r"(\d{2,15})", id)[0]
            employee['employee_'+str(index)]["id"] = int(id_)
        except:
            employee['employee_'+str(index)]["id"] = ''
        try:    
            employee['employee_'+str(index)]["active_date"] = employee['employee_'+str(index)]["active_date"].strftime('%d-%m-%Y')
        except:
            try:
                date_obj = datetime.strptime(employee['employee_'+str(index)]["active_date"], '%m/%d/%Y')
                formatted_date = date_obj.strftime('%d-%m-%Y')
                employee['employee_'+str(index)]["active_date"] = formatted_date
            except:
                employee['employee_'+str(index)]["active_date"] = ''
        try:
            employee['employee_'+str(index)]["assigned_date"] = employee['employee_'+str(index)]["assigned_date"].strftime('%d-%m-%Y')
        except:
            try:
                employee['employee_'+str(index)]["assigned_date"] = datetime.strptime(employee['employee_'+str(index)]["assigned_date"], '%m/%d/%Y').strftime('%d-%m-%Y')
            except:
                employee['employee_'+str(index)]["assigned_date"] = ''

        if pd.isna(employee['employee_'+str(index)]["rate"]):
            employee['employee_'+str(index)][f"rate"] = ''
        else:
            try:
                employee['employee_'+str(index)][f"rate"] = round(employee['employee_'+str(index)]["rate"], 2)
            except:
                if employee['employee_'+str(index)][f"rate"].startswith('='):
                    employee['employee_'+str(index)][f"rate"] = round(float(employee['employee_'+str(index)]["rate"][1:]), 2)
                    
    
        try:
            if employee['employee_'+str(index)]["prorate_amount"].startswith('='):
                expression = employee['employee_'+str(index)]["prorate_amount"][1:]
                workable_days = expression.split('/')[1].split('*')[0]
                employee['employee_'+str(index)]["prorate_amount"] = round(eval(f"{employee['employee_'+str(index)]['rate']}/{workable_days}*{employee['employee_'+str(index)]['credit_days']}"), 2)
        except:
            try:
                employee['employee_'+str(index)]["prorate_amount"] = round(employee['employee_'+str(index)]["prorate_amount"], 2)
            except:
                employee['employee_'+str(index)]["prorate_amount"] = ''

        actual_index = employee['employee_'+str(index)]["rate_adj_date"]
        try:
            date1 = actual_index.iloc[0].strftime('%m-%d-%Y')
        except:
            try:
                date_obj = datetime.strptime(actual_index, '%m/%d/%Y')
                date1 = date_obj.strftime('%m-%d-%Y')
            except:
                try:
                    date_obj = datetime.strptime(str(actual_index).split()[0], '%Y-%m-%d')
                    date1 = date_obj.strftime('%m-%d-%Y')
                except:
                    date1 = ''
        try:
            date2 = actual_index.iloc[1].strftime('%m-%d-%Y')
        except:
            date2 = ''
        try:
            date3 = actual_index.iloc[2].strftime('%m-%d-%Y')
        except:
            date3 = ''
        try:
            date4 = actual_index.iloc[3].strftime('%m-%d-%Y')
        except:
            date4 = ''
        dates = [date1, date2, date3, date4]
        new_dates = [date for date in dates if date]  # Remove empty strings
        today = datetime.now().date()
        try:
            near = min(new_dates, key=lambda x: abs(datetime.strptime(x, '%m-%d-%Y').date() - today))
            employee['employee_'+str(index)]["rate_adj_date"] = near
        except:
            employee['employee_'+str(index)]["rate_adj_date"] = ''
            
        if employee['employee_'+str(index)]["equipment_rental_fee"] == None:
            employee['employee_'+str(index)]["equipment_rental_fee"] = 0
        index += 1
        employees.append(employee)
    return employees
