import openpyxl
import os

my_direct = os.path.dirname(os.path.abspath(__file__))
workbook = openpyxl.load_workbook('{}\prices.xlsx'.format(my_direct))

def read():
    worksheet = workbook.get_sheet_by_name('Sheet1')

    worksheet.iter_cols(min_row=1, min_col=1)
    worksheet['A1']

print(1)