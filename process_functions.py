#!/usr/bin/env python

import openpyxl
import unittest

def count_rows(worksheet):
  row = 1
  while (worksheet['A' + str(row)].value != None):
    row += 1
  return row - 1

class TestProcessFunctions(unittest.TestCase):
  def setUp(self):
    wb = openpyxl.load_workbook('test1973.xlsx', guess_types=True)
    self.ws = wb.active
  
  def test_count_rows(self):
    self.assertEqual(129, count_rows(self.ws))

if __name__ == '__main__':
  unittest.main()
