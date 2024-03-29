#! python3

import os
import sys
from dotenv import load_dotenv
import openpyxl
import requests
import json
import re

if len(sys.argv) < 2:
    print('Usage: python3 geocode.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

load_dotenv()
apikey = os.environ.get("APIKEY")

headers = {
    "apikey": apikey
}

def geocode(address):
    print(address)
    params = (
       ("text","%s, United States"% address),
    )
    response = requests.get('https://app.geocodeapi.io/api/v1/search', headers=headers, params=params)
    if response.status_code != 200:
        coords = ' '
    else:
        parsed_data = json.loads(response.text)
        if len(parsed_data['features']) >= 1:
            coords = parsed_data['features'][0]['geometry']['coordinates']
        else:
            coords = ' '
    return coords

print("Finding coordinates:")
for sheet in workbook.worksheets:
    titlerow = sheet[1]
    state = sheet.title
    print("State: %s"% state)
    addresscol = -1
    coordinatecol = -1
    count = 0
    for cell in titlerow:
        if cell.value == 'Address':
            addresscol = count
        if cell.value == 'Coordinates':
            coordinatecol = count
        count += 1
    if addresscol == -1:
        break
    if coordinatecol == -1:
        sheet.insert_cols(8)
        coordinatecol = 7
        sheet['H1'].value = 'Coordinates'
    for row in sheet.iter_rows(min_row=2):
        # convert to coord
        address = row[addresscol].value
        if address:
            if not address.startswith(" "):
                pattern = "\d.+[A-Za-z]+,\s[A-Z]{2}\s\d{5}-?\d{0,4}.*"
                if re.match(pattern, address):
                    san_pat = "\d.+[A-Za-z]+,\s[A-Z]{2}\s\d{5}-?\d{0,4}"
                    sanitized_address = (re.search(san_pat, address)).group(0)
                    coordinates = str(geocode(sanitized_address))
                    print(coordinates)
                    row[coordinatecol].value = coordinates

# write to coordinatecol
print("Writing file...")
split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-coordinates' + split_filename[1]
workbook.save(new_filename)
print("Done.")
