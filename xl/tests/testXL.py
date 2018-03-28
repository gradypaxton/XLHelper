"""
Test xlutils
"""
import sys
sys.path.insert(0, '..')

import xlutils

print("\nTesting of the XLHelper capabilities\n")
print dir(xlutils)
print

xl = xlutils.XLUtil("test1")

sheets = xl.get_sheets()
print(sheets)

active = xl.get_active_sheet()
print(active)

newSheet = "New Sheet"

xl.make_sheet(newSheet)
print(xl.get_sheets())

xl.select_sheet(newSheet)
print(xl.get_active_sheet())

xl.remove_sheet(newSheet)
print(xl.get_sheets())

sheet = xl.get_sheets()[0]
xl.select_sheet(sheet)

xl.write(1, 1, "Hello World")
print(xl.read(1,1))


xl.write_row(1, 2, ['A', 'B', 'C'])
print(xl.read_row(1, 2, 3))

xl.write_column(1, 3, ['X', 'Y', 'Z'])
print(xl.read_column(1, 3, 3))

xl.style(1, 1, font=xlutils.FONT_BOLD, align=xlutils.ALIGN_CENTER,
         fill=xlutils.FILL_GREY)

xl.save_workbook()
