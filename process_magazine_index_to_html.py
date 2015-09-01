#!/usr/bin/env python

import openpyxl
import process_functions

#want at least 4 html outputs
#chronological listing of every issue
#chrono listing of every article
#listing of every article in every category
#listing of every article by every author

#also want to create a new folder of PDFs with a tiny amount of 
#obfuscation in place to prevent someone from guessing every PDF name 
#and downloading them all (want to at least give people a challenge).

wb = openpyxl.load_workbook('index.xlsx', guess_types=True)
ws = wb.active;

rows = process_functions.count_rows(ws)
print ws['A2250'].value
print rows
