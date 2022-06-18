import xlrd

wb = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\LCC.xls')
sheet_heating = wb.sheet_by_name('heating results')
sheet_elec = wb.sheet_by_name('elec result')

heating_combi_COP_1 = []
heating_results_COP_1 = []
# elec_combi = []
# elec_results = []
# PV_system = []
# g = []


for i in range(sheet_heating.nrows-1):

    heating_combi_COP_1.append(str(sheet_heating.cell_value(i+1,0))+'_'+str(sheet_heating.cell_value(i+1,1))+'_'+str(sheet_heating.cell_value(i+1,2))+'_'+str(sheet_heating.cell_value(i+1,3)))
    heating_results_COP_1.append(float(sheet_heating.cell_value(i+1,4)))

    # if str(sheet_heating.cell_value(i+1, 2)) == 'HCOP-1.0':
    #     heating_combi_COP_1.append(str(sheet_heating.cell_value(i+1,0))+'_'+str(sheet_heating.cell_value(i+1,1))+'_'+str(sheet_heating.cell_value(i+1,2)))
    #     heating_results_COP_1.append(float(sheet_heating.cell_value(i+1,4)))
    # else:
    #     heating_combi_COP_3.append(str(sheet_heating.cell_value(i+1,0))+'_'+str(sheet_heating.cell_value(i+1,1))+'_'+str(sheet_heating.cell_value(i+1,2))+'_'+str(sheet_heating.cell_value(i+1,3)))
    #     heating_results_COP_3.append(float(sheet_heating.cell_value(i+1,4)))


for j in  range(sheet_elec.nrows-1):
    elec_combi = (str(sheet_elec.cell_value(j + 1, 0)) + '_' + str(sheet_elec.cell_value(j + 1, 1)) + '_' + str(sheet_elec.cell_value(j + 1, 2))+ '_' + str(sheet_elec.cell_value(j + 1, 4)))
    elec_results = (str(sheet_elec.cell_value(j + 1, 5)))
    PV_system = (str(sheet_elec.cell_value(j + 1, 3)))
    g = (str(sheet_elec.cell_value(j + 1, 4)))
    heating_index = heating_combi_COP_1.index(elec_combi)
    LCC_totoal = float(elec_results) + float(heating_results_COP_1[heating_index])
    print(str(sheet_elec.cell_value(j + 1, 0)) + '_' + str(sheet_elec.cell_value(j + 1, 1)) + '_' + str(sheet_elec.cell_value(j + 1, 2)) + "_" + PV_system + '_'+ g+ "_" + str(LCC_totoal))
    # if str(sheet_elec.cell_value(j+1, 2)) == 'HCOP-1.0':
    #     elec_combi = (str(sheet_elec.cell_value(j+1,0))+'_'+str(sheet_elec.cell_value(j+1,1))+'_'+str(sheet_elec.cell_value(j+1,2)))
    #     elec_results = (str(sheet_elec.cell_value(j+1, 5)))
    #     PV_system = (str(sheet_elec.cell_value(j+1,4)))
    #     g = (str(sheet_elec.cell_value(j+1,3)))
    #     heating_index = heating_combi_COP_1.index(elec_combi)
    #     LCC_totoal = float(elec_results) + float(heating_results_COP_1[heating_index])
    #     print(elec_combi + "_" + PV_system + "_" + g + "_" + str(LCC_totoal))
    # else:
    #     elec_combi1 = (str(sheet_elec.cell_value(j+1,0))+'_'+str(sheet_elec.cell_value(j+1,1))+'_'+str(sheet_elec.cell_value(j+1,2))+'_'+str(sheet_elec.cell_value(j+1,3)))
    #     elec_results1 = (str(sheet_elec.cell_value(j+1, 5)))
    #     PV_system = (str(sheet_elec.cell_value(j+1,4)))
    #     g = (str(sheet_elec.cell_value(j+1,3)))
    #     heating_index = heating_combi_COP_3.index(elec_combi1)
    #     LCC_totoal = float(elec_results1) + float(heating_results_COP_3[heating_index])
    #     print(str(sheet_elec.cell_value(j+1,0))+'_'+str(sheet_elec.cell_value(j+1,1))+'_'+str(sheet_elec.cell_value(j+1,2)) + "_" + PV_system + "_" + g + "_" + str(LCC_totoal))


