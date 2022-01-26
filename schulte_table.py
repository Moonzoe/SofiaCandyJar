"""
Script to generate 5*5 Schulte Table base on an excel template
"""
import pathlib
import random
import numpy as np
import openpyxl

EXCEL_TEMPLATE = pathlib.Path(__file__).parent / 'assets' / 'schulte_template.xlsx'
START_COLUMN = 2
DELTA_COLUMN = 6
LINE_ONE_START_ROW = 2
LINE_TWO_START_ROW = 12


def generate_single_schulte_table(level=5):
    size_of_schulte = level
    li = [i for i in range(1, size_of_schulte * size_of_schulte + 1)]
    random.shuffle(li)
    arr = np.reshape(li, (size_of_schulte, size_of_schulte))
    return arr


def write_excel():
    file = openpyxl.load_workbook(EXCEL_TEMPLATE)
    sheet = file.worksheets[0]
    arr_list = [generate_single_schulte_table(5) for i in range(6)]
    for square_no, square_values in enumerate(arr_list):
        start_column = START_COLUMN
        if square_no < 3:
            start_row = LINE_ONE_START_ROW
        else:
            start_row = LINE_TWO_START_ROW
        for row_index, row in enumerate(square_values):
            for column_index, cell_value in enumerate(row):
                sheet.cell(row=row_index + start_row,
                           column=column_index + start_column + DELTA_COLUMN * (square_no % 3)).value = cell_value
    file.save('the_schulte_table.xlsx')


if __name__ == '__main__':
    write_excel()
