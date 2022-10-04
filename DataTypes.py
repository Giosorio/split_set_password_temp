import pandas as pd
import datetime


def textpercent_to_number(x):
    """Remove the percent (%) from the string and convert it into a float value with only 3 decimals"""

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
    """Transform datetime values into string dates"""

    if isinstance(x, datetime.datetime):
        x = x.strftime('%d/%m/%Y')
    
    return x


def reduce_decimals(x):
    """Float values are rounded to only 3 decimals
    i.e If there is a formula on the excel file it will take all the decimals, this function reduces the decimals"""

    if type(x) is float:
        x = round(x,3)
    
    return x


def transform_special_columns(df, file):

    special_cols_dict = {
        textpercent_to_number: ['UK Holiday Pay ST', 'UK Pension Auto Enrollment ST', 'UK Employer NI ST', 'UK Apprenticeship Levy ST', 'Supplier markup % on standard Pay Rate', 'Standard Supplier Bill Rate'], 
        date_string: ["Worker Date of Birth", "Current Assignment Original Start Date", "Current Assignment End Date"],
        reduce_decimals: ['Supplier markup % on standard Pay Rate']
    }

    for func_apply, special_cols in special_cols_dict.items():
        for c in special_cols:
            if c in df.columns:
                df[c] = df[c].apply(func_apply)
            # else:
            #     print(c, 'not in template', file)

    return df



