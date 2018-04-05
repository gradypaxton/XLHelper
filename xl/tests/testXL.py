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

print(xl.get_coord('A1'))
print(xl.get_coord('AA12'))

print(xl.make_coord(1,1))
print(xl.make_coord(27,12))

xl.write("a1", "Testing")
print(xl.read('a1'))


xl.write_row('a2', ['A', 'B', 'C'])
print(xl.read_row('a2', 3))

xl.write_column('a3', ['X', 'Y', 'Z'])
print(xl.read_column('a3', 3))

xl.style(1, 1, font=xlutils.FONT_BOLD, align=xlutils.ALIGN_CENTER,
         fill=xlutils.FILL_GREY)


xl.write_row('c4', ['ALPHA', 'BETA', 'GAMMA'])
xl.style_row(3, 4, 3, font=xlutils.FONT_BOLD,
             align=xlutils.ALIGN_CENTER, fill=xlutils.FILL_GREY)

xl.write_column('g3', ['CHI', 'PSI', 'OMEGA'])
xl.style_column(7, 3, 3, font=xlutils.FONT_BOLD,
             align=xlutils.ALIGN_CENTER, fill=xlutils.FILL_GREY)

xl.freeze(2, 2)

xl.set_column_width(1, auto=True)
xl.set_column_width(2, w=25)


print xl.save_workbook()
