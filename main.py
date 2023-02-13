from glob import glob
import pandas as pd

year = '2023'
months = '02'

months = months.replace(' ', '').split(',')

source_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - ' \
      'Hospital/Preregistration/Daily Reports/'
dest_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - ' \
       'Hospital/Preregistration/Daily Reports/Consolidated Files/'

for month in months:
    data = []
    criteria = f'{source_dir}*{year}-{month}*.xlsx'
    files = glob(criteria)
    for file in files:
        try:
            df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts')
        except ValueError:
            df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts - Consolidated')
        data.append(df)

    data = pd.concat(data)
    data.to_excel(f'{dest_dir}{year} {month} Combined.xlsx', index = False)
    print(f'{month} {year} saved')
