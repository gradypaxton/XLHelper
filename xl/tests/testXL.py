"""
Test xlutils
"""
import sys
sys.path.insert(0, '..')

import xlutils

print dir(xlutils)
xl = xlutils.XLUtil("test1")
xl.write(1, 1, "Hello World")
xl.save_workbook()
