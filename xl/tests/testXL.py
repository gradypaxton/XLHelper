"""
Test xlutils
"""
import sys
sys.path.insert(0, '..')

import xlutils

print("\nTesting of the XLHelper capabilities\n")
print dir(xlutils)
print

print "XLUtil Docs"
print xlutils.XLUtil.__doc__

print "Methods"
print dir(xlutils.XLUtil)
print

xl = xlutils.XLUtil("test1")
print xl.__init__.__doc__


print xl.get_sheets.__doc__
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


print(xl.make_coord(1, 1))
print(xl.make_coord(27, 12))


print(xl.get_span('A1:b2'))
print(xl.make_span([1, 1], [2, 2]))


xl.write("a1", "Testing")
print(xl.read('a1'))


xl.write_row('a2', ['A', 'B', 'C'])
print(xl.read_row('a2', 3))


xl.write_column('a3', ['X', 'Y', 'Z'])
print(xl.read_column('a3', 3))


xl.write_block('a10', [ ['ACDC', 'BTO'],
                        ['Align Tech', 'Boeing', 'Citigroup'],
                        ['Audi', 'Buick']])
print(xl.read_block('a10:c12'))


xl.style('a1', font=xlutils.FONT_BOLD, align=xlutils.ALIGN_CENTER,
         fill=xlutils.FILL_GREY)



xl.write_row('c4', ['ALPHA', 'BETA', 'GAMMA'])
xl.style_block('c4:e4', font=xlutils.FONT_BOLD,
             align=xlutils.ALIGN_CENTER, fill=xlutils.FILL_GREY)

xl.write_column('g3', ['CHI', 'PSI', 'OMEGA'])
xl.style_block('g3:g5', font=xlutils.FONT_BOLD,
             align=xlutils.ALIGN_CENTER, fill=xlutils.FILL_GREY)

xl.freeze('b2')

xl.set_column_width(1, auto=True)
xl.set_column_width(2, w=25)


print xl.save_workbook()
