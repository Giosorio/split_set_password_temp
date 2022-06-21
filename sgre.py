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


def create_output_folders(today):
    folder_xl_files = f'xl_files-{today}'
    folder_xl_files_password = f'xl_files_password-{today}'
    path_xl_files = f'{path}{folder_xl_files}\\'
    path_xl_files_password = f'{path}{folder_xl_files_password}\\'
    
    try:
        os.mkdir(folder_xl_files)
        os.mkdir(folder_xl_files_password)
    except FileExistsError:
        pass

    return path_xl_files, path_xl_files_password


def sort_df_passwords(df, today):

    # First_letter_upper
    df['col_'] = [supp.upper() for supp in df['Supplier']]
    df.sort_values(by='col_', inplace=True)
    df = df.drop(columns=['col_'])
    df.to_excel(f'{today} - Siemens_passwords.xlsx', index=False)


def start():

    today = datetime.today().strftime('%Y%m%d')
    path_xl_files, path_xl_files_password = create_output_folders(today)
    
    df = pd.read_excel(path+file, sheet_name='report', header=None, na_filter=False)
    
    header = df.iloc[1]
    df = df.iloc[2:, :]

    # df[9] = pd.to_datetime(df[9], format='%Y-%m-%d %H:%M:%S')
    date_cols = [9,10,15]
    for d_col in date_cols:
        df[d_col] = pd.to_datetime(df[d_col], infer_datetime_format=True)
        df[d_col] = df[d_col].dt.strftime('%d/%m/%Y')
        print(df[d_col])

    df.columns = header

    suppliers = set(df['Supplier'])
    print(suppliers)
    print(len(suppliers))


    password_master = []
    for supp in suppliers:
        name = supp.replace('/', '').replace('[','').replace(']','').replace('|','')
        df_supp = df[df['Supplier']==supp]
        file_name = f'SIE1 - {name} - {today}.xlsx'
        with pd.ExcelWriter(f'{path_xl_files}{file_name}', engine='xlsxwriter') as writer:
            df_supp.to_excel(writer, sheet_name='Sheet1', index=False) 

            wb = writer.book
            ws = writer.sheets['Sheet1']

            hd_format = wb.add_format({'text_wrap':True, 'bold':True})

            for i, col in enumerate(df.columns):
                ws.set_column_pixels(i, i, width=150)
                ws.write(0, i, col, hd_format)

        pw = create_password()
        PassProtect(f'{path_xl_files}{file_name}', pw, f'{path_xl_files_password}{file_name}')
        password_master.append((supp, pw))

    
    df_pw = pd.DataFrame(password_master, columns=['Supplier', 'Password'])
    sort_df_passwords(df_pw, today)



if __name__ == '__main__':
    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\split_set_password\\'
    file = 'SE_Security ID Master.xlsx'

    start()

    
   