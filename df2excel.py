import xlsxwriter 
from datetime import datetime

import win32com.client as win32


def excel_col_w_fitting(excel_path, sheet_name_list):
    """
    This function make all column widths of an Excel file auto-fit with the column content.
    :param excel_path: The Excel file's path.
    :param sheet_name_list: The sheet names of the Excel file.
    :return: File's column width correctly formatted.
    """

    import win32com.client as win32

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    work_book = excel.Workbooks.Open(excel_path)

    for sheet_name in sheet_name_list:
        work_sheet = work_book.Worksheets(sheet_name)
        work_sheet.Columns.AutoFit()
    work_book.Save()
    excel.Application.Quit()
    return None


def df_format(df_data_list, sheet_name_list, start_row_list, start_col_list, save_path):
    """
    This function format a list of DataFrames and save them to their worksheets on the target Excel file.
    :param df_data_list: list of DataFrames.
    :param sheet_name_list: list of worksheet names.
    :param start_row_list: the start row of every sheet.
    :param start_col_list: the start column of every sheet.
    :param save_path: the saving path of the generated workbook.
    :return: Saved file.
    """

    import xlsxwriter
    from datetime import datetime

    # Create a workbook.
    work_book = xlsxwriter.Workbook(save_path)

    # Define header format.
    header_format = work_book.add_format()
    header_format.set_font('Arial')
    header_format.set_font_size(10)
    header_format.set_bold()
    header_format.set_align('center')
    header_format.set_align('vcenter')
    header_format.set_border()

    # Define norm data format.
    data_format = work_book.add_format()
    data_format.set_font('Arial')
    data_format.set_font_size(10)
    data_format.set_border()

    # Define datetime format.
    dt_format = work_book.add_format()
    dt_format.set_font('Arial')
    dt_format.set_font_size(10)
    dt_format.set_num_format('yyyy/mm/dd h:mm:ss')
    dt_format.set_border()

    for i, sheet_name in enumerate(sheet_name_list):
        df_data, start_row, start_col = df_data_list[i], start_row_list[i], start_col_list[i]

        # Add a worksheet.
        work_sheet = work_book.add_worksheet(sheet_name)

        # Write header.
        header = df_data.columns.tolist()
        for j, col in enumerate(header):
            work_sheet.write(start_row, start_col + j, col, header_format)

        t_sample = datetime(2020, 1, 1, 0, 0, 0)
        # Write data.
        for j in range(df_data.shape[1]):
            temp = df_data.iloc[0, j]
            temp_out = 0
            # noinspection PyBroadException
            try:
                if isinstance(temp.date(), type(t_sample.date())):
                    temp_out = 'datetime'
            except Exception:
                temp_out = 'not datetime'

            for k in range(df_data.shape[0]):
                if temp_out == 'datetime':
                    work_sheet.write(start_row + 1 + k, start_col + j, dt_format)
                else:
                    work_sheet.write(start_row + 1 + k, start_col + j, data_format)

    work_book.close()

    excel_col_w_fitting(save_path, sheet_name_list)

    return None
