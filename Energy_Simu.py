import xlwt
import xlrd
#Read from result
workbook = xlrd.open_workbook(r"C:\Users\Adam\Desktop\AEBN20\Energy_simu\Basecase1.xls")
booksheet = workbook.sheet_by_name('Basecase1')
#writing
write_workbook = xlwt.Workbook(encoding='utf-8')
booksheet02 = write_workbook.add_sheet('Basecase1', cell_overwrite_ok=True)
col = 0
col = 0

for row in range(booksheet.nrows):
    row_content = booksheet.row_values(row)