# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the delimiter.
# An example output can be seen in the "example.csv" file.
import xlrd
import os
import csv
from zipfile import ZipFile
datafile = "../2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    
    # For each column, need [name, date, val]
    data = None
    
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    data = []
    for col in range(1, sheet.ncols):
        station = sheet.cell_value(0, col)
        if station is "ERCOT": continue
        max_load, date = get_max_val_and_date(sheet, col)
        row = {
            "Station":  station,
            "Year":     date[0],
            "Month":    date[1],
            "Day":      date[2],
            "Hour":     date[3],
            "Max Load": max_load
        }
        data.append(row)    
    return data

def get_max_val_and_date(sheet, col):
    coast_vals     = sheet.col_values(col, 1)
    maxvalue       = max(coast_vals)
    maxpos         = coast_vals.index(maxvalue) + 1
    max_time_excel = sheet.cell_value(maxpos, 0)
    max_time       = xlrd.xldate_as_tuple(max_time_excel, 0)
    return (maxvalue, max_time)

def save_file(data, filename):
    with open(filename, "wb") as f:
        r = csv.DictWriter(f, ["Station","Year","Month","Day","Hour","Max Load"], delimiter="|")
        r.writeheader()
        r.writerows(data)

def test():
    open_zip(datafile)
    data = parse_file(datafile)
    save_file(data, outfile)

    ans = {'FAR_WEST': {'Max Load': "2281.2722140000024", 'Year': "2013", "Month": "6", "Day": "26", "Hour": "17"}}
    
    fields = ["Year", "Month", "Day", "Hour", "Max Load"]
    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            s = line["Station"]
            if s == 'FAR_WEST':
                for field in fields:
                    assert ans[s][field] == line[field]


