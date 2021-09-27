#! python3

import os
import sys
import openpyxl
import requests
import json

if len(sys.argv) < 2:
    print('Usage: python3 geocode.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

headers = {
  "apikey": os.getenv('apikey')
}

def geocode(address,city,state):
    params = (
       ("text","%s, %s, %s, United States"% address, city, state),
    );
    response = requests.get('https://app.geocodeapi.io/api/v1/search', headers=headers, params=params)
    parsed_data = json.loads(response.text)
    coords = parsed_data['features'][0]['geometry']['coordinates']
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
        coordinates = geocode(address,city,state)
        row[coordinatecol] = coordinates

# write to coordinatecol
split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-coordinates' + split_filename[1]
workbook.save(new_filename)
