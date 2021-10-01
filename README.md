# excel-geocoder
A python script that takes an excel document as input, recognizes a column labeled "Address", and creates a new column labeled "Coordinates". The addresses are then converted into geographic coordinates and added to the "Coordinates" column. Writes data to a new excel file "FileName-geocoded.xlsx" with added "Coordinates" column.

Requires a mapbox-gl apikey.

Usage: python3 geocode.py filename.xlsx
