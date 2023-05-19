from glob import glob
import pandas as pd
from tkinter import *
from tqdm import tqdm
import os


def combine(year, month):
    root.destroy()
    
    source_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/Preregistration/Daily Reports/'
    dest_dir = 'C:/Users/dpashayan/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/Preregistration/Daily Reports/Consolidated Files/'
    
    if not os.path.exists(source_dir):
        print("Please be sure you are connected to the VPN and have the sharepoint folder accessible through Windows File Explorer")
    else:
        data = []
        criteria = f'{source_dir}*{year}-{month}*.xlsx'
        files = glob(criteria)
        for file in tqdm(files):
            try:
                df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts')
            except ValueError:
                df = pd.read_excel(file, engine='openpyxl', sheet_name='Accounts - Consolidated')
            data.append(df)

        data = pd.concat(data)
        if not os.path.exists(dest_dir):
            os.mkdir(dest_dir)
        data.to_excel(f'{dest_dir}{year} {month} Combined.xlsx', index = False)
        print(f'{month} {year} saved')

if __name__ == '__main__':
    root = Tk()

    year = StringVar()
    month = StringVar()

    Label(root, text = 'Enter the year in YYYY format').pack(padx = 5, pady=5)
    Entry(root, textvariable = year).pack(padx = 5, pady=5)
    Label(root, text = 'Enter the month in MM format').pack(padx = 5, pady=5)
    Entry(root, textvariable = month).pack(padx = 5, pady=5)
    Button(root, text = 'Combine', command = lambda: combine(year.get(), month.get())).pack(padx = 5, pady=5)

    root.mainloop()