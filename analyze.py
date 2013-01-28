#!/usr/bin/env python

import sys
import xlrd
import datetime

def main():
    print sys.argv[1]
    workbook = xlrd.open_workbook(sys.argv[1])

    worksheets = workbook.sheet_names()
    #for worksheet_name in worksheets:

    worksheet = workbook.sheet_by_name('1A')
    print worksheet

    nrows = worksheet.nrows
    for r in xrange(1, nrows):
        #print worksheet.row(r)
        meal_num = int(worksheet.cell_value(r, 9))
        start_time = xlrd.xldate_as_tuple(worksheet.cell_value(r, 4), 1)
        end_time = xlrd.xldate_as_tuple(worksheet.cell_value(r, 6), 1)
        diff_time = xlrd.xldate_as_tuple(worksheet.cell_value(r, 6) - worksheet.cell_value(r, 4), 1)
        print r, meal_num, start_time, end_time, diff_time

if __name__ == '__main__':
  main()
