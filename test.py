import pandas as pd
# import win32com.client as win32  ## pip install pywin32
# from win32com.client.gencache import EnsureDispatch


def test1():

    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\split_set_password\\'
    path_out = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\xl\\xlfiles\\files_pw\\'
    file = 'date_test.xlsx'

    df = pd.read_excel(path+file)
    print(df['date_col'].head())   

    df['date_col'] =  df['date_col'].dt.strftime('%d/%m/%Y')
    print(df['date_col'].head())   


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


if __name__ == '__main__':

    print('adfadsf')










