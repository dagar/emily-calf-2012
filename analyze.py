#!/usr/bin/env python
''' Emily Miller-Cushon 2013 '''

import sys
import xlrd
import datetime

YEAR = 2012
MONTH = 06

def convert_excel_date(xldate, day):
    date = xlrd.xldate_as_tuple(xldate, 1)
    return datetime.datetime(YEAR, MONTH, day, date[3], date[4], date[5])

def read_sheet(xls_name, sheet_name):
    ''' reads each sheet from the excel file '''
    workbook = xlrd.open_workbook(xls_name)
    worksheet = workbook.sheet_by_name(sheet_name)

    calf = {}

    nrows = worksheet.nrows
    meal_start_time = worksheet.cell_value(1, 4)
    for current_row in xrange(1, nrows):
        week = int(worksheet.cell_value(current_row, 2))
        day = int(worksheet.cell_value(current_row, 3))
        end_time = worksheet.cell_value(current_row, 6)
        meal_num = int(worksheet.cell_value(current_row, 9))

        if current_row < nrows - 1:
            next_meal_num = int(worksheet.cell_value(current_row + 1, 9))
        else:
            next_meal_num = -1

        if not calf.has_key(day):
            calf[day] = []

        if next_meal_num != meal_num:
            calf[day].append([convert_excel_date(meal_start_time, day), convert_excel_date(end_time, day)])

            #print worksheet.cell_value(current_row, 0), convert_excel_date(meal_start_time, day), convert_excel_date(end_time, day)
            if next_meal_num == -1:
                return calf
            else:
                meal_start_time = worksheet.cell_value(current_row + 1, 4)

def overlap_time(start1, end1, start2, end2):
    ''' compute overlaping time between two intervals '''

    if end1 < start2 or end2 < start1:
        return 0
    else:
        #print start1, end1
        #print start2, end2
        times = [start1, end1, start2, end2]
        times.sort()
        return times[2] - times[1]


def compare_calves(pen, calfa, calfb):

    day_sum = {}
    for day in calfa.keys():
        day_sum[day] = 0
        for intervala in calfa[day]:
            for intervalb in calfb[day]:
                current_overlap = overlap_time(intervala[0], intervala[1], intervalb[0], intervalb[1])
                if (current_overlap != 0):
                    day_sum[day] += current_overlap.seconds
        print "%d, %d, %d" % (pen, day, day_sum[day])


    return day_sum


def main():
    xls_name = sys.argv[1]

    print "Pen, Day, Overlap"
    for pen in range(1, 11):
        calfa = read_sheet(xls_name, '%dA' % pen)
        calfb = read_sheet(xls_name, '%dB' % pen)

        compare_calves(pen, calfa, calfb)



if __name__ == '__main__':
    main()
