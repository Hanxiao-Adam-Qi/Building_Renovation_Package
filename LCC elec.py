import xlrd

wb_elec = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\PV_savings.xls')
sheet_elec_savings_SEK = wb_elec.sheet_by_name('SEK')
sheet_elec_use_annually = wb_elec.sheet_by_name('elec_use_annually')

wb_price = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\LCC.xls')
sheet_general_price = wb_price.sheet_by_name('LCC inputs')

wb_elec_use_hourly = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\elec_use_hourly.xls')
sheet_elec_use_hourly = wb_elec_use_hourly.sheet_by_name('electricity use hourly')

elec_annnually_use_title = []
elec_annnually_use_list = []
for u in range(sheet_elec_use_hourly.ncols-1):
    elec_use_hourly = []
    for v in  range(sheet_elec_use_hourly.nrows-1):
        elec_use_hourly.append(sheet_elec_use_hourly.cell_value(v+1, u+1))
    elec_use_annually = sum(elec_use_hourly)
    # print('{}'.format(sheet_elec_use_hourly.cell_value(0, u+1))+"_"+str(elec_use_annually))
    elec_annnually_use_list.append(elec_use_annually)
for z in range(sheet_elec_use_annually.nrows-1):
    elec_annnually_use_title.append(str(sheet_elec_use_annually.cell_value(z+1, 0))+"_"+str(sheet_elec_use_annually.cell_value(z+1, 1))+"_"+str(sheet_elec_use_annually.cell_value(z+1, 2)))

g_EL = [-0.0042, 0.005, 0.02]
i_EL = float(sheet_general_price.cell_value(6, 5))
yr = int(sheet_general_price.cell_value(12, 5))

PV_names = ['Large system', 'Small system', 'Big system']
PV_cost = [133866, 77854, 184162]

for g in g_EL:
    PV_system = []
    elec_savings = []

    insul_type = []
    win_type = []
    heat_pump = []

    for i in range(sheet_elec_savings_SEK.nrows-1):
        PV_system.append(sheet_elec_savings_SEK.cell_value(i+1, 3))
        elec_savings.append(sheet_elec_savings_SEK.cell_value(i+1, 4))

        insul_type.append(sheet_elec_savings_SEK.cell_value(i+1, 0))
        win_type.append(sheet_elec_savings_SEK.cell_value(i+1, 1))
        heat_pump.append(sheet_elec_savings_SEK.cell_value(i+1, 2))

    for j in range(len(elec_savings)):



        PV_index =PV_names.index(PV_system[j])
        initial_cost = PV_cost[PV_index]

        elec_annnually_use_index = elec_annnually_use_title.index(str(insul_type[j]+"_"+str(win_type[j]+"_"+str(heat_pump[j]))))
        elec_use = elec_annnually_use_list[elec_annnually_use_index]

        # LCC_elec = initial_cost + ((elec_use*(1+g)) * (1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g)) - ((elec_savings[j]*(1+g)) * (1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g))

        LCC_elec = initial_cost + ( - ((elec_savings[j]*(1+g)) * (1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g)))

        print('{}_{}_{}_{}_{}'.format(insul_type[j], win_type[j], heat_pump[j], PV_system[j], g)+'_'+str(LCC_elec))
        # print(LCC_elec)

        # paste the results in LCC.xls-elec results