import pandas as pd
import xlsxwriter


def header_format(wb, ws, df, format_header_row, header_index, example_hd_index=None, example_hd_format=None, example_index=None, example_format=None):
    """
    Sets up the header format
    header_index -> index in the dataframe where the header is located    
    example_hd_index -> index in the dataframe where the example header is located  
    example_index -> index in the dataframe where the row example is located   
    """

    # hd_format1 = wb.add_format({'bg_color': '#008080', 'bold':1, 'font_color':'#FFFFFF', 'text_wrap':True})
    
    def get_format(bg_color, font_color='#000000', bold=True, italic=False, text_wrap=True):
        """
        bg_color : back ground color
        font_color: default black
        bold: default true
        italic: default false
        text_wrap: default true
        """
        new_format = wb.add_format({'bg_color': bg_color, 'bold':bold, 'italic':italic, 'font_color':font_color, 'text_wrap':text_wrap}) 
        return new_format

    bg_blue_font_black = get_format(bg_color='#BBCCE4')
    bg_green_font_black = get_format(bg_color='#D8E4BC')
    bg_orange_font_black = get_format(bg_color='#FCD5B4')

    ## Personal information
    bg_grey_font_white = get_format(bg_color='#6F8693', font_color='#FFFFFF')  # personal_info
    ## ACTIVE WORKER?
    bg_light_purple_font_white = get_format(bg_color='#912CEE', font_color='#FFFFFF')  # active_wk
    ## Supplier fields
    bg_dark_purple_font_white = get_format(bg_color='#4B0082', font_color='#FFFFFF')  # supplier_fields
    ## Standard Rates
    bg_dark_blue_font_white = get_format(bg_color='#366092', font_color='#FFFFFF')  # st_rate
    bg_red_font_white = get_format(bg_color='#CD0000', font_color='#FFFFFF')  # us_rates
    ## Other rates
    bg_dark_green_font_white = get_format(bg_color='#308014', font_color='#FFFFFF')  # shift_2
    bg_dark_yellow_font_white = get_format(bg_color='#8B7500', font_color='#FFFFFF') # shift_3
    ## Comms
    bg_orange_font_white = get_format(bg_color='#F79646', font_color='#FFFFFF')  # comms

    bg_white_font_black_nw = get_format(bg_color='#FFFFFF', text_wrap=False)  # example_hd_f
    bg_light_blue_font_black_nw = get_format(bg_color='#ADD8E6', bold=False, italic=True, text_wrap=False)  # example_row
    italic_wrap = get_format(bg_color='#FFFFFF', italic=True, bold=False)

    for col_num, col_format in zip(df.columns, format_header_row):
        string_header = df.iloc[header_index, col_num]
        ws.write(header_index, col_num, string_header, eval(col_format))


    if not example_hd_index is None:
        for col_num in df.columns:
            string_header = df.iloc[example_hd_index, col_num]
            ws.write(example_hd_index, col_num, string_header, eval(example_hd_format))
    

    if not example_index is None:
        for col_num in df.columns:
            string_header = df.iloc[example_index, col_num]
            ws.write(example_index, col_num, string_header, eval(example_format))