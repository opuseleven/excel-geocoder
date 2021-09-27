#! python3

import os
import sys
import openpyxl

if len(sys.argv) < 2:
    print('Usage: python3 geocode.py filename.xl')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

addresscol = null
coordinatecol = null

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    for cell in titlerow:
        if cell = 'Address':
            addresscol = cell.column
        if cell = 'Coordinates':
            coordinatecol = cell.column
    if addresscol = null:
        sys.exit()
    if coordinatecol = null:
        sheet.insert_cols(8)
        coordinatecol = sheet['H']
    for cell in addresscol:
        print(cell.value)
        # convert to coord
        # write to coordinatecol
