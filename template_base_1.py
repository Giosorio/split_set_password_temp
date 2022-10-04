import pandas as pd
import xlsxwriter
import os
from os import listdir, mkdir
from os.path import isfile, join
import glob
import string
import random
import platform
from datetime import datetime

######## WIN32 ONLY WORKS IN WINDOWS
# import win32com.client as win32  ## pip install pywin32
# from win32com.client.gencache import EnsureDispatch



def PassProtect_win32(read_path, pw, new_path=None):
    """win32 ONLY WORKS IN WINDOWS"""
    
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


def create_password(supplier, random_pw=False):
    """If random_pw is False is because there will be multiple batches and the password must remain the same"""
    
    if random_pw is True:
        letters = string.ascii_uppercase + string.digits
        random_6 =  ''.join(random.choice(letters) for _ in range(6))
        password = project_name + random_6
        
        return password

    supplier = ''.join(char for char in supplier if char.isalnum())
    l = len(supplier)
    pw = project_name + str(123 * l) + supplier[:3][::-1]  # supplier[:-4:-1]
    return pw.upper()


def check_encrypt_method():

    op_system = platform.system()

    if set_password_method == 'msoffice-crypt':
        return 'msoffice-crypt'
    elif set_password_method == 'win_32' & op_system != 'Windows':
        print('Password cannot be set to the files, win32 only works in Windows \nUse "msoffice-crypt" as a method to encrypt the files')
        return None
    elif set_password_method == 'win_32' & op_system == 'Windows':
        return 'win_32'


def set_password(path_1, path_2, passwordMaster_name, set_password_method='msoffice-crypt'):

    def encrypt_file(password, path_in, path_out):
        """msoffice-crypt must be installed in the local folder"""

        os.system(f'msoffice/bin/msoffice-crypt.exe -e -p {password} {path_in} {path_out}')

    df_pw = pd.read_csv(passwordMaster_name)
    num_files = df_pw.shape[0]
    encrypt_method = check_encrypt_method()
    
    count = 1
    for file_n, pw in zip(df_pw['Filename'], df_pw['Password']):
        if encrypt_method == None:
            return None
        
        if encrypt_method == 'win_32':
            PassProtect_win32(f'{path_1}/{file_n}', pw, f'{path_2}/{file_n}')
        
        if encrypt_method == 'msoffice-crypt':              
            path_in = '"{}/{}"'.format(path_1, file_n)
            path_out = '"{}/{}"'.format(path_2, file_n)
            encrypt_file(pw, path_in, path_out)
            print(file_n, '   Password:', pw, f'---->{count}/{num_files}')
            count +=1
    
        
def main():

    today = datetime.now().strftime('%Y%m%d')

    path_1 = f'{project_name}_XL_files_{today}'
    path_2 = f'{project_name}XL_files_pw_{today}'
    os.mkdir(path_1)
    os.mkdir(path_2)

    df = pd.read_excel(file, sheet_name='report', header=None)
    header = df.iloc[1]
    
    df = df.iloc[2:, :]
    df.columns = header


    # df['Worker Date of Birth'] = pd.to_datetime(df['Worker Date of Birth'], format='%d/%m/%Y')
    # df['Worker Date of Birth'] = df['Worker Date of Birth'].dt.strftime('%d/%m/%Y')

    suppliers = set(df['Supplier'])
    print(suppliers)
    print(len(suppliers))


    password_master = []
    for i, supp in enumerate(suppliers,1):
        # remove special characters from the supplier name
        name = ''.join(char for char in supp if char == ' ' or char.isalnum())
        id_file = f'{project_name}ID{batch}{i:03d}'

        df_supp = df[df['Supplier']==supp]
        file_name = f'{id_file}-{name}-{today}.xlsx'
        file_path = f'{path_1}/{file_name}'

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df_supp.to_excel(writer, sheet_name='Sheet1', index=False) 

            wb = writer.book
            ws = writer.sheets['Sheet1']

            hd_format = wb.add_format({'text_wrap':True, 'bold':True})

            for i, col in enumerate(df.columns):
                if i == 15:
                    ws.set_column_pixels(i, i, width=400)
                else:
                    ws.set_column_pixels(i, i, width=150)
                ws.write(0, i, col, hd_format)

        pw = create_password(supp)    
        password_master.append((id_file, file_name, supp, pw))

    df_pw = pd.DataFrame(password_master, columns=['File ID', 'Filename', 'Supplier', 'Password'])
    passwordMaster_name = f'PasswordMaster-{today}.csv'
    df_pw.to_csv(passwordMaster_name, index=False)

    set_password(path_1, path_2, passwordMaster_name, set_password_method)
    


if __name__ == '__main__':
    
    file = 'SIE1SE_Security ID_29092022.xlsx'

    project_name = 'SIE1SE'
    batch = 1
    set_password_method = 'msoffice-crypt'    # 'msoffice-crypt' or 'win_32'

    main()

