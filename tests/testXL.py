
import xlutils

xl = XLUtil("test1")
xl.write(1, 1, "Hello World")
xl.save_workbook()
