"""
This is is used to convert xls files to csv.
Author = Vinay Mavi(vinaymavi@gmail.com)
"""

import xlrd
import csv
import unicodecsv
from os import listdir
# csv files directory path.
PATH = '/Users/vinaymavi/xls-to-csv/files/xls/2015June09/'
dir_list = listdir(PATH)
total = 0
count = 0
for f in dir_list:
    count += 1
    workbook = xlrd.open_workbook(PATH + f)
    sheet = workbook.sheet_by_index(0)
    # csv file path should be exists in path.
    csv_file = open('/Users/vinaymavi/xls-to-csv/files/csv/csv2015June09.csv', 'a')
    wr = unicodecsv.writer(csv_file, quoting=csv.QUOTE_ALL, delimiter=';')
    total += sheet.nrows
    print("Total Rows=", f, count, sheet.nrows)
    try:
        for row_num in xrange(sheet.nrows):
            wr.writerow(sheet.row_values(row_num))
    finally:
        csv_file.close()

print ("Total banks=", total)
