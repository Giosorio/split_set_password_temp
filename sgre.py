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
from HeaderFormat import header_format
from ProtectSheet import lock_only_header


def create_password():
    letters = string.ascii_uppercase + string.digits
    random_10 =  ''.join(random.choice(letters) for i in range(4))
    password = project_name + random_10 + '2022'
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


def column_width(ws, df, width_format=None):
    """
    Sets up the columns width
    index_col_width_dict = Dictionary containing the columns index as keys and the values are the width of the column

    Parameters:
    ws -> worksheet
    df -> dataframe used to create the template header=None
    """
    if width_format is None:
        ws.set_column_pixels(col_num, col_num, width=150)
        return 'Done'
    
    for col_num, width in zip(df.columns, width_format):
        ws.set_column_pixels(col_num, col_num, width=width)


def set_data_validation(ws, df, data_val_row, initial_index):
    """
    initial_index -> index from which the data is based, EXCLUDING THE HEADER (assuming the header willl be locked)
    """

    length = df.shape[0] 
    for col_num, data_val in zip(df.columns, data_val_row):
        if data_val != '':
            data_validation = data_val.split('~')
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            ws.data_validation(f'{col_letter}{initial_index+1}:{col_letter}{length}', {'validate':'list', 'source':data_validation, 'error_type':'warning'})


def start():

    today = datetime.today().strftime('%Y%m%d')
    path_xl_files, path_xl_files_password = create_output_folders(today)
    
    df = pd.read_excel(path+file, sheet_name='temp', header=None, na_filter=False)
    
    width_format = df.iloc[0]
    format_header_row = df.iloc[1]
    data_val_row = df.iloc[2]
    df_header = df.iloc[3:6]
    df = df.iloc[6:]
    initial_index = excel_index_data - 1 # index df starts from 0

    # date_cols = [9,10,15]
    # for d_col in date_cols:
    #     df[d_col] = pd.to_datetime(df[d_col], infer_datetime_format=True)
    #     df[d_col] = df[d_col].dt.strftime('%d/%m/%Y')
    #     print(df[d_col])


    suppliers = set(df[9])
    print(suppliers)
    print(len(suppliers))


    password_master = []
    for i, supp in enumerate(suppliers,1):
        name = ''.join(char for char in supp if char == ' ' or char.isalnum())
        df_supp = df[df['Supplier']==supp]
        id_file = f'{project_name}ID{batch}{i:03d}'
        file_name = f'{project_name} Worker data collection-{id_file}-{name}-{today}.xlsx'

        df_supp = pd.concat([df_header, df_supp], axis=0)
        
        with pd.ExcelWriter(f'{path_xl_files}{file_name}', engine='xlsxwriter') as writer:
            df_supp.to_excel(writer, sheet_name='Sheet1', index=False) 

            wb = writer.book
            ws = writer.sheets['Sheet1']

            header_format(wb, ws, df, format_header_row, header_index=2, example_hd_index=1, example_hd_format='italic_wrap')
            set_data_validation(ws, df_supp, data_val_row, initial_index)

            column_width(ws, df, width_format)

            merge_format1 = wb.add_format({'bg_color': '#D8E4BC', 'bold':1, 'font_color':'#000000'}) 
            merge_format2 = wb.add_format({'bg_color': '#FCD5B4', 'bold':1, 'font_color':'#000000'}) 
            ws.merge_range('AF1:AH1', 'Please fill if Booking Entity belongs to the UK', merge_format1)
            ws.merge_range('AI1:AU1', 'Please fill if Booking Entity belongs to the UK', merge_format2)
            
            lock_only_header(wb, ws, df, initial_index, sheet_password='RSR123')


        pw = create_password()
        PassProtect(f'{path_xl_files}{file_name}', pw, f'{path_xl_files_password}{file_name}')
        password_master.append((supp, pw))

    
    df_pw = pd.DataFrame(password_master, columns=['Supplier', 'Password'])
    sort_df_passwords(df_pw, today)



if __name__ == '__main__':
    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\split_set_password\\'
    file = 'BCOM_wk_data_collection_supplier.xlsx'
    project_name = 'BCOM'
    batch = 1
    excel_index_data = 4 # index where the data starts in the final template (not including the template config rows)

    start()

    
   