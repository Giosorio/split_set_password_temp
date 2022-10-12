from re import X
import pandas as pd
import datetime



def textpercent_to_number(x):
    """Remove the percent (%) from the string and convert it into a float value with only 3 decimals, the function is applied to each value of the column"""

    try:
        y = x.replace('%', '')
        y = float(y)/100
        y = round(y,3)
        return y
    except AttributeError:
        pass
    except ValueError:
        pass
    return x


def date_string(x):
    """Transform datetime values into string dates, the function is applied to each value of the column"""

    if isinstance(x, datetime.datetime):
        x = x.strftime('%d/%m/%Y')
    
    return x


def xlnumber_date(df, col):
    """Transform excel numbers to dates, the function is applied to the entire column at once"""
    
    df[col] = pd.to_datetime(df[col], unit='D', origin=pd.Timestamp('30/12/1899')).dt.strftime('%d/%m/%Y')
    
    return df


def xlnumber_date_each(excel_value):
    """Transform excel numbers to dates, the function is applied to each value of the column"""

    x = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + excel_value - 2).strftime('%d/%m/%Y')

    return x


def reduce_decimals(x):
    """Float values are rounded to only 3 decimals, the function is applied to each value of the column
    i.e If there is a formula on the excel file it will take all the decimals, this function reduces the decimals"""

    if type(x) is float:
        x = round(x,3)
    
    return x


def transform_special_columns(df, special_cols_dict):

    functions_to_apply_to_entire_column = ['xlnumber_date']

    for func_apply, special_cols in special_cols_dict.items():
        for col in special_cols:
            if func_apply in functions_to_apply_to_entire_column:
                print(func_apply, col)
                df = eval(func_apply)(df, col)
            elif col in df.columns:
                df[col] = df[col].apply(eval(func_apply))
            # else:
            #     print(col, 'not in template')

    return df



