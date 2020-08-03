import xlsxwriter
import datetime

# xlsxwriter tutorial: https://xlsxwriter.readthedocs.io/index.html

# Create a workbook and add a worksheet.
work_book = xlsxwriter.Workbook('Expenses03.xlsx')
work_sheet = work_book.add_worksheet()

# Add a bold format to use to highlight cells.
bold = work_book.add_format()
bold.set_bold()

# Add a number format for cells with money.
money_format = work_book.add_format({'num_format': '$#,##0'})

# Add an Excel date format.
date_format = work_book.add_format({'num_format': 'yyyy/mm/dd'})

# Adjust the column width.
work_sheet.set_column(1, 1, 15)

# Write some data headers.
work_sheet.write('A1', 'Item', bold)
work_sheet.write('B1', 'Date', bold)
work_sheet.write('C1', 'Cost', bold)

# Some data we want to write to the worksheet.
expenses = (
    ['Rent', '2013-01-13', 1000],
    ['Gas', '2013-01-14', 100],
    ['Food', '2013-01-16', 300],
    ['Gym', '2013-01-20', 50]
)

# Start from the first cell.
row, col = 1, 0

# Iterate over the data and write it out row by row
for item, date_str, cost in expenses:
    # Convert the date string into a datetime object.
    date = datetime.datetime.strptime(date_str, '%Y-%m-%d')

    work_sheet.write(row, col, item)
    work_sheet.write(row, col + 1, date, date_format)
    work_sheet.write(row, col + 2, cost, money_format)
    row += 1

# Write a total using a formula.
work_sheet.write(row, 0, 'Total', bold)
work_sheet.write(row, 2, '=SUM(B1:B4)', money_format)

work_book.close()
