import openpyxl
import os
import numpy


class Excel_py:
    my_direct = os.path.dirname(os.path.abspath(__file__))
    workbook = openpyxl.load_workbook('{}\prices.xlsx'.format(my_direct))

    def read(self, min_row=1, min_col=1, worksheet_name='A1'):
        worksheet = self.workbook.get_sheet_by_name('Sheet1')
        worksheet.iter_cols(min_row=min_row, min_col=min_col)

        return worksheet[worksheet_name]

    def get_data(self):
        data = {
            'A': [self.read(worksheet_name='A{}'.format(index+1)).value for index in range(251)],
            'B': [self.read(worksheet_name='B{}'.format(index+1)).value for index in range(251)],
            'C': [self.read(worksheet_name='C{}'.format(index+1)).value for index in range(251)],
            'D': [self.read(worksheet_name='D{}'.format(index+1)).value for index in range(251)],
            'E': [self.read(worksheet_name='E{}'.format(index+1)).value for index in range(251)],
            'F': [self.read(worksheet_name='F{}'.format(index+1)).value for index in range(251)],
            'G': [self.read(worksheet_name='G{}'.format(index+1)).value for index in range(251)],
            'H': [self.read(worksheet_name='H{}'.format(index+1)).value for index in range(251)],
            'I': [self.read(worksheet_name='I}'.format(index+1)).value for index in range(251)],
            'J': [self.read(worksheet_name='J{}'.format(index+1)).value for index in range(251)],
            'K': [self.read(worksheet_name='K{}'.format(index+1)).value for index in range(251)],
            'L': [self.read(worksheet_name='L{}'.format(index+1)).value for index in range(251)],
        }

        return data


# = numpy.mean()
