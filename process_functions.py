#!/usr/bin/env python

def count_rows(worksheet):
  row = 1
  while (worksheet['A' + str(row)].value != None):
    row += 1
  return row
