#!/usr/bin/env python

import openpyxl
import unittest
import string

def count_rows(worksheet):
  row = 1
  while (worksheet['A' + str(row)].value != None):
    row += 1
  return row - 1

def convert_colnum_to_char(column):
  return string.ascii_uppercase[column - 1]

def count_cols(worksheet):
  col = 1
  while (worksheet[convert_colnum_to_char(col) + '1'].value != None):
    col += 1
  return col - 1

def make_dict(worksheet):
  rows = count_rows(worksheet)
  cols = count_cols(worksheet)

class TestProcessFunctions(unittest.TestCase):
  def setUp(self):
    wb = openpyxl.load_workbook('test1973.xlsx', guess_types=True)
    self.ws = wb.active
    self.rowCount = count_rows(self.ws)
    self.colCount = count_cols(self.ws)
  
  def test_count_rows(self):
    self.assertEqual(129, self.rowCount)
    
  def test_colnum_char_converter(self):
    self.assertEqual('A', convert_colnum_to_char(1))
    self.assertEqual('M', convert_colnum_to_char(13))
    self.assertEqual('Z', convert_colnum_to_char(26))

  def test_count_cols(self):
    self.assertEqual(11, self.colCount)

if __name__ == '__main__':
  unittest.main()
