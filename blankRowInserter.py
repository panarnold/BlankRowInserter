#! python
# blankRoWInserter.py - script responsible for adding blank rows to the existing excel file
# run onto terminal: python blankRowInserter.py <M> <N> random_name.xlsx - and it
# creates new excel file (based on random_name.xlsx), with new M rows, started from <N> row of existing worksheet
# X 2020 Arnold Cytrowski

import openpyxl, sys
from openpyxl import workbook

from openpyxl.descriptors.base import Integer

if len(sys.argv) != 4:
    print('usage: python blankRowInserter.py <M> <N> <excel_filename>')
    sys.exit(-1)

m = int(abs(sys.argv[1]))
n = int(abs(sys.argv[2]))



workbook = None
try:
    workbook = openpyxl.open(sys.argv[3])
except:
    print('sorry, that wasn\'t possible, try again')
    sys.exit(-1)

sheet = workbook.active

sheet.insert_rows(m, n)

workbook.save(f'updated{sys.argv[3]}')
print('aaand it\'s done')
sys.exit(0)



