#!/usr/bin/env python
''' Emily Miller-Cushon 2013 '''

import sys
import xlrd
import datetime

def convert_excel_date(xldate, day):
    date = xlrd.xldate_as_tuple(xldate, 1)

    return datetime.datetime(2012, 06, day, date[3], date[4], date[5])


def read_sheet(xls_name, sheet_name):
    ''' reads each sheet from the excel file '''
    workbook = xlrd.open_workbook(xls_name)
    worksheet = workbook.sheet_by_name(sheet_name)

    ret = []

    nrows = worksheet.nrows
    meal_start_time = worksheet.cell_value(1, 4)
    for current_row in xrange(1, nrows):
        calf = worksheet.cell_value(current_row, 0)
        week = int(worksheet.cell_value(current_row, 2))
        day = int(worksheet.cell_value(current_row, 3))
        end_time = worksheet.cell_value(current_row, 6)
        meal_num = int(worksheet.cell_value(current_row, 9))

        if current_row < nrows - 1:
            next_meal_num = int(worksheet.cell_value(current_row + 1, 9))
        else:
            next_meal_num = -1

        if next_meal_num != meal_num:
            ret.append([calf, week, day, convert_excel_date(meal_start_time, day), convert_excel_date(end_time, day)])
            if next_meal_num == -1:
                return ret
            else:
                meal_start_time = worksheet.cell_value(current_row + 1, 4)

def overlap_time(start1, end1, start2, end2):
    ''' compute overlaping time between two intervals '''
    if end1 < start2 or end2 < start1:
        return 0
    else:
        return max(end1, end2) - min(start1, start2)

def main():
    xls_name = sys.argv[1]
    calf_a = read_sheet(xls_name, '1A')
    calf_b = read_sheet(xls_name, '1B')


if __name__ == '__main__':
    main()
