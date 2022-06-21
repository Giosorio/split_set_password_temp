import pandas as pd
import xlsxwriter


class ProtectSheet:

    # locked = wb.add_format({'locked': True})
    unlocked_text = wb.add_format({'locked': False, 'text_wrap':True})
    unlocked_Amount = wb.add_format({'locked': False, 'text_wrap':False, 'num_format': f'{currency}#,##0.00'})
    locked_Amount = wb.add_format({'text_wrap':False, 'num_format': f'{currency}#,##0.00'})    
    unlocked_percent = wb.add_format({'locked': False, 'text_wrap':False, 'num_format': '0.00%'})

    @classmethod
    def unlock_only_data(wb, ws, df, initial_index, sheet_password, unlocked_row, currency='£'):
        """
        Lock the entire sheet except the data range from the dataframe 

        Sets up the format of each column in the dataframe 
        unlocked_row -> List containing the format per each column of the df, the format is applied only to the range of the data frame NOT THE ENTIRE COLUMN
        initial_index -> index from which the data is based, EXCLUDING THE HEADER (assuming the header willl be locked)
        """
        
        for col, unlocked_f in zip(df.columns, unlocked_row):
            unlocked_cells = df.iloc[initial_index:, col]
            if unlocked_f != '':
                ws.write_column(initial_index, col, unlocked_cells, cell_format=eval(unlocked_f))

        ws.protect(sheet_password)

    @classmethod
    def lock_only_header(wb, ws, df, initial_index, sheet_password, currency='£', unlocked_row=None, last_index_unlocked=1000):


        width = df.shape[1]
        last_column_letter = xlsxwriter.utility.xl_col_to_name(width-1)

        if not unlocked_row is None:
            ws.unprotect_range(f'A{initial_index+1}:{last_column_letter}{last_index_unlocked}')
            return 'Done'
        
        for col in df.columns:
            col_letter = xlsxwriter.utility.xl_col_to_name(col)
            ws.unprotect_range(f'{col_letter}{initial_index+1}:{col_letter}{last_index_unlocked}')
        
        ws.protect(sheet_password)
    
    