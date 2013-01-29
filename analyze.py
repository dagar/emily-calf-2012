#!/usr/bin/env python

import sys
import xlrd
import datetime

def convert_excel_date(xldate, day):
    date = xlrd.xldate_as_tuple(xldate, 1)

    return datetime.datetime(2012, 06, day, date[3], date[4], date[5])


def read_sheet(xls_name, sheet_name):
    workbook = xlrd.open_workbook(xls_name)
    worksheet = workbook.sheet_by_name(sheet_name)

    ret = []

    nrows = worksheet.nrows
    current_meal_num = int(worksheet.cell_value(1, 9))
    meal_start_time = worksheet.cell_value(1, 4)
    for r in xrange(1, nrows):
        calf = worksheet.cell_value(r, 0)
        trmt = worksheet.cell_value(r, 1)
        week = int(worksheet.cell_value(r, 2))
        day = int(worksheet.cell_value(r, 3))
        start_time = worksheet.cell_value(r, 4)
        scan = worksheet.cell_value(r, 5)
        end_time = worksheet.cell_value(r, 6)

        meal_num = int(worksheet.cell_value(r, 9))

        if r < nrows - 1:
            next_meal_num = int(worksheet.cell_value(r + 1, 9))
        else:
            next_meal_num = -1

        if next_meal_num != meal_num:
            feeding_time = end_time - meal_start_time
            ret.append([calf, week, day, convert_excel_date(meal_start_time, day), convert_excel_date(end_time, day), convert_excel_date(feeding_time, day)])
            current_meal_num = meal_num
            if next_meal_num == -1:
                return ret
            else:
                meal_start_time = worksheet.cell_value(r + 1, 4)


def main():
    workbook = read_sheet(sys.argv[1], '1A')

    print workbook


if __name__ == '__main__':
    main()
