#!/usr/bin/env python3

import xlrd
import csv
import sys

import pprint as pp

xlsx_file  = sys.argv[1]
xlsx_sheet = sys.argv[2]
csv_file   = sys.argv[3]

def read_excel(xlsx_file, xlsx_sheet):

    wb    = xlrd.open_workbook(xlsx_file)
    sheet = wb.sheet_by_name(xlsx_sheet)

    return sheet

def write_as_cv(sheet, csv_file):

    #your_csv_file = open(csv_file, 'wb')
    your_csv_file = open(csv_file, 'w')

    #wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_MINIMAL)

    for rownum in range(sheet.nrows):
        wr.writerow(sheet.row_values(rownum))
        #row_values=[]
        #for value in sheet.row_values(rownum):
        #    if "," in value and value[0] == '"' and value[ len(value)-1 ] == '"':
        #        print("Changing value from <" + value + ">")
        #        value = value[1:-1]
        #        print("To                  <" + value + ">")
        # 
        #    row_values.append(value)
        #wr.writerow(row_values)

    your_csv_file.close()

sheet = read_excel(xlsx_file, xlsx_sheet)

pp.pprint(sheet)

write_as_cv(sheet, csv_file)


