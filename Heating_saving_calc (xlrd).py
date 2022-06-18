import xlrd
import xlwt

# Calculate heating saving kWh annually.

wb_heating_use_hourly = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\Heating hourly use.xls')
sheet_heating_use_hourly = wb_heating_use_hourly.sheet_by_name('Heating use hourly')

for i in range(sheet_heating_use_hourly.ncols-1):
    saving_hourly = []
    initial_hourly = []
    for j in range(sheet_heating_use_hourly.nrows-1):
        initial_hourly.append(sheet_heating_use_hourly.cell_value(j+1, i+1))
        saving_hourly_1 = sheet_heating_use_hourly.cell_value(j+1, 1) - sheet_heating_use_hourly.cell_value(j+1, i+1)
        saving_hourly.append(saving_hourly_1)
    saving_annually = sum(saving_hourly)
    initial_use_annually = sum(initial_hourly)
    # print('{}_'.format(sheet_heating_use_hourly.cell_value(0,i+1)) + str(saving_annually))
    print(initial_use_annually)
