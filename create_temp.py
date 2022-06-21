import pandas as pd
import xlsxwriter
import win32com.client as win32  ## pip install pywin32
from win32com.client.gencache import EnsureDispatch
import os
from os import listdir
from os.path import isfile, join
import glob
import string
import random
from datetime import datetime


def create_password():
    letters = string.ascii_uppercase + string.digits
    random_10 =  ''.join(random.choice(letters) for i in range(4))
    password = 'SIE1' + random_10 + '2022'
    # print(password)
    return password


def PassProtect(read_path, pw, new_path=None):
    
    xlApp = EnsureDispatch("Excel.Application") 
    xlwb = xlApp.Workbooks.Open(read_path)
    xlApp.DisplayAlerts = False
    xlwb.Visible = True
    
    if new_path != None:
        xlwb.SaveAs(new_path, Password = pw)
    else:
        xlwb.SaveAs(read_path, Password = pw)
    
    xlwb.Close()
    xlApp.Quit()


def column_width():
    pass


if __name__ == '__main__':
    
    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\xl\\xlfiles\\'
    path_out = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\xl\\xlfiles\\files_pw\\'
    file = '_UK__IAM__Master_Worker_Record (08.04.2022).xlsx'

    df = pd.read_excel(path+file, sheet_name='report', header=None)
    header = df.iloc[1]
    
    df = df.iloc[2:, :]
    df.columns = header

    df_filter = df[(df['JS - Security ID collected?']==0) & (df['WK - Security ID collected?']==0)]
    
    df_filter['JS - Security ID collected?'] = ''
    df_filter['WK - Security ID collected?'] = ''
    print(df_filter.dtypes)
    print(df_filter['Worker Date of Birth'])
    df_filter['Worker Date of Birth'] = pd.to_datetime(df_filter['Worker Date of Birth'], format='%d/%m/%Y')
    print(df_filter['Worker Date of Birth'])
    df_filter['Worker Date of Birth'] = df_filter['Worker Date of Birth'].dt.strftime('%d/%m/%Y')
    print(df_filter['Worker Date of Birth'])
    df_filter = df_filter.iloc[:, :-2]
    df_filter.to_csv('filter_496.csv', index=False)

    suppliers = set(df_filter['Supplier'])
    print(suppliers)
    print(len(suppliers))


    password_master = []
    for supp in suppliers:
        name = supp.replace('/', '').replace('[','').replace(']','')
        df_supp = df_filter[df_filter['Supplier']==supp]
        file_name = f'SIE1 - {name} - 2022.xlsx'
        with pd.ExcelWriter(f'{path}{file_name}', engine='xlsxwriter') as writer:
            df_supp.to_excel(writer, sheet_name='Sheet1', index=False) 

            wb = writer.book
            ws = writer.sheets['Sheet1']

            hd_format = wb.add_format({'text_wrap':True, 'bold':True})

            for i, col in enumerate(df_filter.columns):
                ws.set_column_pixels(i, i, width=150)
                ws.write(0, i, col, hd_format)

        pw = create_password()
        PassProtect(f'{path}{file_name}', pw, f'{path_out}{file_name}')
        password_master.append((supp, pw))

    
    df_pw = pd.DataFrame(password_master, columns=['Supplier', 'Password'])
    df_pw.to_csv('Siemens_passwords.csv', index=False)