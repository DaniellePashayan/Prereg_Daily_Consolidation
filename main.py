import pandas as pd
from glob import glob


def main():
    year = input('Enter the year you wish to combine in YYYY format: ')
    months = input('Enter the month you wish to combine in MM format. If using more than one month, separate using comma: ')
    months = months.replace(' ', '').split(',')

    source_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - ' \
          'Hospital/Preregistration/Daily Reports/'
    dest_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - ' \
           'Hospital/Preregistration/Daily Reports/Consolidated Files/'

    for month in months:
        data=[]
        criteria = f'{source_dir}*{year}-{month}*.xlsx'
        files = glob(criteria)
        for file in files:
            try:
                df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts')
            except ValueError:
                df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts - Consolidated')
            data.append(df)

        data = pd.concat(data)
        data.to_excel(f'{dest_dir}{year} {month} Combined.xlsx', index =None)
        print(f'{month} {year} saved')


if __name__ == '__main__':
    main()