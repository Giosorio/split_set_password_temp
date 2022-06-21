import pandas as pd


def test1():

    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\split_set_password\\'
    path_out = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\xl\\xlfiles\\files_pw\\'
    file = 'date_test.xlsx'

    df = pd.read_excel(path+file)
    print(df['date_col'].head())   

    df['date_col'] =  df['date_col'].dt.strftime('%d/%m/%Y')
    print(df['date_col'].head())   


if __name__ == '__main__':
    test1()
    















