#! python3

import os
import sys
import openpyxl

if len(sys.argv) < 2:
    print('Usage: python3 geocode.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

def geocode(address):
    coords = ""
    return coords

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    state = sheet.title
    addresscol = null
    coordinatecol = null
    for cell in titlerow:
        if cell.value = 'Address':
            addresscol = cell.column
        if cell.value = 'Coordinates':
            coordinatecol = cell.column
    if addresscol = null:
        break
    if coordinatecol = null:
        sheet.insert_cols(8)
        coordinatecol = sheet['H']
        sheet['H1'] = 'Coordinates'
    for row in sheet.iter_rows(min_row=2):
        # convert to coord
        address = row[addresscol].value
        city = row[citycol].value
        coordinates = geocode(address)
        row[coordinatecol] = coordinates

# write to coordinatecol
split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-coordinates' + split_filename[1]
workbook.save(new_filename)
