import pandas as pd
from fuzzywuzzy import fuzz

# Cargar los datos de los dos archivos Excel en DataFrames
excel1 = pd.read_excel('output.xlsx')
excel2 = pd.read_excel('output2.xlsx')

errors_log = {
    "logs": {}
}

employees_not_found = []

# Iterar sobre los ID en el primer archivo Excel y verificar si están presentes en el segundo
for id in excel1['id']:
    if id in excel2['id'].values:
        
        empleado_excel1 = excel1[excel1['id'] == id]
        empleado_excel2 = excel2[excel2['id'] == id]
        

        # Creamos un diccionario para almacenar las diferencias para este empleado
        differences = {}

        # Comparamos las columnas para encontrar las diferencias
        for column in excel1.columns:
            value_excel1 = empleado_excel1[column].values[0]
            value_excel2 = empleado_excel2[column].values[0]
            if column == 'id':
                continue
            elif column == 'company_name':
                # value_excel1 = empleado_excel1[column].values[0]
                differences['company_name'] = value_excel1
                continue
            elif column == 'file_name':
                differences['file_name_SDM'] = value_excel1
                differences['file_name_IA'] = value_excel2
                continue
            elif column == 'status':
                differences['status_IA'] = value_excel2
                pass
            try:
                # Aplicar .strip() solo si el valor es una cadena
                if isinstance(value_excel1, str):
                    value_excel1 = value_excel1.strip()
                if isinstance(value_excel2, str):
                    value_excel2 = value_excel2.strip()
            except AttributeError:
                # Ignorar el error si el valor no se puede convertir a cadena
                pass
            if value_excel1 != value_excel2:
                if pd.isna(value_excel1) and pd.isna(value_excel2):
                    continue
                else:
                    if pd.isna(value_excel1):
                        value_excel1 = 'Empty'
                    elif pd.isna(value_excel2):
                        value_excel2 = 'Empty'
                    if value_excel1 == 'Empty' and value_excel2 == 0.0:
                        continue
                    elif column == 'name':
                        name_sdm = value_excel1
                        name_ia = value_excel2
                        ratio = fuzz.ratio(name_sdm.lower(), name_ia.lower())
                        if ratio > 70:
                            continue
                        else:
                            differences[column] = f'value SDM: {value_excel1}, value IA: {value_excel2}'
                    differences[column] = f'value SDM: {value_excel1}, value IA: {value_excel2}'

        # Si hay diferencias para este empleado, las agregamos al registro de errores
        if differences:
            errors_log["logs"][int(id)] = differences
    else:
        empleado_excel1 = excel1[excel1['id'] == id]
        try:
            employee = {
                'id':id,
                'company_name':empleado_excel1['company_name'].values[0],
                'status':empleado_excel1['status'].values[0]
            }
        except:continue
        employees_not_found.append(employee.values())
# Crear un DataFrame para la hoja "summary_differences"
df = pd.DataFrame()

# Iterar sobre las claves (IDs) en errors_log['logs']
for id, data in errors_log['logs'].items():
    # Iterar sobre las columnas incorrectas para este ID
    for columna, valor in data.items():
        # Si la columna ya existe en el DataFrame, agregamos un nuevo valor
        # de lo contrario, la creamos y establecemos los valores anteriores como NaN
        if columna in df.columns:
            df.at[id, columna] = valor
        else:
            df.at[id, columna] = valor

# Crear DataFrames para las hojas "rates" y "rate adjustment date"
rates_data = []
rate_adj_date_data = []
title_data = []

# Iterar sobre las claves (IDs) en errors_log['logs']
for id, data in errors_log['logs'].items():
    # Verificar si hay diferencias en las columnas "rate" y "rate adjustment date" para este empleado
    if 'rate' in data:
        # Agregar valores a la hoja "rates"
        rates_data.append(excel1.loc[excel1['id'] == id, ['id', 'company_name', 'rate']].values[0])
        # Agregar valores a la hoja "rate adjustment date"
for id, data in errors_log['logs'].items():
    if 'rate_adj_date' in data:
        rate_adj_date_data.append(excel1.loc[excel1['id'] == id, ['id', 'company_name', 'rate_adj_date']].values[0])
for id, data in errors_log['logs'].items():
    if 'title' in data:
        title_data.append(excel1.loc[excel1['id'] == id, ['id', 'company_name', 'title']].values[0])

# Convertir listas de datos en DataFrames
rates_df = pd.DataFrame(rates_data, columns=['id', 'company_name', 'rate'])
rate_adj_date_df = pd.DataFrame(rate_adj_date_data, columns=['id', 'company_name', 'rate_adj_date'])
title_df = pd.DataFrame(title_data, columns=['id', 'company_name', 'title'])
employees_df = pd.DataFrame(employees_not_found, columns=['id', 'company_name', 'status'])

# Guardar los DataFrames en un nuevo archivo Excel
with pd.ExcelWriter('summary_differences.xlsx') as writer:
    df.to_excel(writer, sheet_name='summary_differences', index_label='ID')
    rates_df.to_excel(writer, sheet_name='rates', index=False)
    rate_adj_date_df.to_excel(writer, sheet_name='rate adjustment date', index=False)
    title_df.to_excel(writer, sheet_name='title', index=False)
    employees_df.to_excel(writer, sheet_name='employees not found', index=False)

print("El archivo 'summary_differences.xlsx' se ha creado exitosamente.")
