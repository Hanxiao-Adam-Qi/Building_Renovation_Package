import xlrd
import xlwt

# LCC for heating part
wb_heating = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\Heating savings.xls')
sheet_heating_kWh = wb_heating.sheet_by_name('kWh')
sheet_heating_prices = wb_heating.sheet_by_name('price')

wb_elec = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\PV_savings.xls')
sheet_elec_use = wb_elec.sheet_by_name('elec_use_annually')

wb_price = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\LCC.xls')
sheet_general_price = wb_price.sheet_by_name('LCC inputs')

g_EL = [-0.0042, 0.005, 0.02]

changing_name = []
initial_cost = []
for i in range(sheet_heating_prices.nrows-1):
    changing_name.append(str(sheet_heating_prices.cell_value(i+1, 0))+ '-' + str(sheet_heating_prices.cell_value(i+1,1)))
    initial_cost.append(sheet_heating_prices.cell_value(i+1, 2))
# print(changing_name)
# print(initial_cost)

insul_type = []
win_type = []
heat_pump = []
heating_combi = []
heating_savings = []
heating_uses = []

insul_type_elec = []
win_type_elec = []
heat_pump_elec = []
elec_combi = []
elec_uses = []

for v in range(sheet_heating_kWh.nrows-1):
    insul_type.append(str(sheet_heating_kWh.cell_value(v+1, 0)))
    win_type.append(str(sheet_heating_kWh.cell_value(v+1, 1)))
    heat_pump.append(str(sheet_heating_kWh.cell_value(v+1, 2)))
    heating_combi.append(str(sheet_heating_kWh.cell_value(v+1, 0))+"_"+str(sheet_heating_kWh.cell_value(v+1, 1))+"_"+str(sheet_heating_kWh.cell_value(v+1, 2)))
    heating_savings.append(float(sheet_heating_kWh.cell_value(v+1, 3)))
    heating_uses.append(float(sheet_heating_kWh.cell_value(v+1, 4)))

for uu in range(sheet_elec_use.nrows-1):
    insul_type_elec.append(str(sheet_elec_use.cell_value(uu+1, 0)))
    win_type_elec.append(str(sheet_elec_use.cell_value(uu+1, 1)))
    heat_pump_elec.append(str(sheet_elec_use.cell_value(uu+1, 2)))
    elec_combi.append(str(sheet_elec_use.cell_value(uu+1, 0))+"_"+str(sheet_elec_use.cell_value(uu+1, 1))+"_"+str(sheet_elec_use.cell_value(uu+1, 2)))
    elec_uses.append((14926/12775.38199)*float(sheet_elec_use.cell_value(uu+1, 3)))    #correct with annually error


# print(insul_type)
# print(win_type)
# print(heat_pump)
# print(heating_savings)

DH_price = sheet_general_price.cell_value(3, 8)
g_DH = sheet_general_price.cell_value(5, 8)
i_DH = sheet_general_price.cell_value(6, 8)
fixed_price_DH = sheet_general_price.cell_value(9, 8)
yr = int(sheet_general_price.cell_value(12, 8))


g_EL = [-0.0042, 0.005, 0.02]
i_EL = float(sheet_general_price.cell_value(6, 5))

# print(DH_price)
# print(g_DH)
# print(i_DH)
# print(fixed_price_DH)
# print(yr)


# Heating LCC calc
for g in g_EL:
    for u in range(len(insul_type)):
        insu_index = changing_name.index(insul_type[u])
        insu_cost = float(initial_cost[insu_index])

        win_index = changing_name.index((win_type[u]))
        win_cost = float(initial_cost[win_index])

        heat_pump_index = changing_name.index(heat_pump[u])
        heat_pump_cost = float(initial_cost[heat_pump_index])

        LCC_DH_initial_cost = insu_cost + win_cost + heat_pump_cost
        LCC_DH_fixed_price = (fixed_price_DH) * ((1+i_DH)**yr-1)/(i_DH*(1+i_DH)**yr)
        LCC_DH_heating_savings = (heating_savings[u]*(1+g_DH))*((1-(1+g_DH)**yr * (1+i_DH)**-yr)/(i_DH-g_DH))
        LCC_DH_heating_uses = (heating_uses[u]*0.854*(1+g_DH))*((1-((1+g_DH)**yr) * ((1+i_DH)**-yr))/(i_DH-g_DH))
        LCC_EL_fixed_price = 7014.792* ((1+i_DH)**yr-1)/(i_DH*(1+i_DH)**yr)
        LCC_DH = LCC_DH_initial_cost + LCC_DH_heating_uses

        elec_comb_index = elec_combi.index(heating_combi[u])
        elec_use_with_heating_combi = elec_uses[elec_comb_index]

        # print('{}_{}_{}'.format(insul_type[u], win_type[u], heat_pump[u])+'_'+str(LCC_DH))


        # print(LCC_DH_heating_uses)
        # print((elec_use_with_heating_combi)*((1+g)*(1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g)))

        if heat_pump[u] == 'HCOP-3.0':
            print('{}_{}_{}_{}'.format(insul_type[u], win_type[u], heat_pump[u], g)+'_'+str(LCC_DH+(elec_use_with_heating_combi*0.966)*((1+g)*(1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g))+ LCC_EL_fixed_price))
        else:
            print('{}_{}_{}_{}'.format(insul_type[u], win_type[u], heat_pump[u], g)+'_'+str(LCC_DH+(elec_use_with_heating_combi*0.966)*((1+g)*(1-(1+g)**yr*(1+i_EL)**-yr)/(i_EL-g)) + LCC_DH_fixed_price+ LCC_EL_fixed_price))

