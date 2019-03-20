import openpyxl
import os

my_direct = os.path.dirname(os.path.abspath(__file__))

workbook = openpyxl.load_workbook('{}\prices.xlsx'.format(my_direct))
worksheet = workbook.get_sheet_by_name('Sheet1')

print(1)