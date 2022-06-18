import xlrd
import xlwt

wb = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\LCC.xls')
sheet_result = wb.sheet_by_name('LCC_total')
sheet_check = wb.sheet_by_name('over prod check')

check_combi = []
check = []
for j in range(sheet_check.nrows-1):
    check_combi.append(sheet_check.cell_value(j+1, 0) +"_"+sheet_check.cell_value(j+1, 1) +"_"+sheet_check.cell_value(j+1, 2) +"_"+sheet_check.cell_value(j+1, 3))
    check.append(sheet_check.cell_value(j+1, 4))

for i in range(sheet_result.nrows-1):
    if sheet_result.cell_value(i+1, 3) == '-':
        a = '-'
    else:
        index = check_combi.index(sheet_result.cell_value(i+1, 0)+"_"+sheet_result.cell_value(i+1, 1)+"_"+sheet_result.cell_value(i+1, 2)+"_"+sheet_result.cell_value(i+1, 3))
        a = check[index]

    print(a)