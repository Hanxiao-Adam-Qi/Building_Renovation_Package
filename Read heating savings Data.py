import xlrd
import xlwt

Insulation_Type = ['Glasswool', 'SIP']
Insulation_Thinkness = [0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3]
Window_Types = [0,1, 2, 3, 4, 5, 6]
HP_COP = [1, 3]

years = 30
DH_price = 0.854
DH_g = 0.01
DH_i = 0.03
DH_ini_fix = 4542.108

# workbook = xlwt.Workbook(encoding = 'utf-8')
# cop1 = workbook.add_sheet('cop=1', cell_overwrite_ok=True)
# cop2 = workbook.add_sheet('cop=2', cell_overwrite_ok=True)



for i in Insulation_Type:
    if i == Insulation_Type[0]:
        for j in Insulation_Thinkness:
            for k in Window_Types:
                for l in HP_COP:
                    wb1 = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\Energy_simu\python_grasshopper data\{}-{}_Window-{}_HCOP-{}.xls'.format(i, j, k, l))
                    obj1 = wb1.sheet_by_index(0)
                    energy_savings1 = sum(obj1.col_values(9,1,13))




                    

                    # energy_savings1_price = DH_price*(1+DH_g)*(1-(1+DH_g)^years*(1+DH_i))/(DH_i-DH_g)
                    # DH_fixed_cost = -DH_ini_fix*((1+DH_i)^years-DH_i*years-1)/((DH_i^2)*(1+DH_i)^years)
                    # Initial_cost =
                    print('{}-{}_Window-{}_HCOP-{}'.format(i, j, k, l),'_', energy_savings1)
                    # print(energy_savings1)
                    # print('\n')


    else:
        for j in Insulation_Thinkness[1:4]:
            for k in Window_Types:
                for l in HP_COP:
                    wb2 = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\Energy_simu\python_grasshopper data\{}-{}_Window-{}_HCOP-{}.xls'.format(i, j, k, l))
                    obj2 = wb2.sheet_by_index(0)
                    energy_savings2 = sum(obj2.col_values(9,1,13))

                    print('{}-{}_Window-{}_HCOP-{}'.format(i, j, k, l), '_', energy_savings2)

                    # print('{}-{}_Window-{}_HCOP-{}.xls'.format(i, j, k, l))
                    # print(energy_savings2)
                    # print('\n')

# workbook.save(r'C:\Users\Adam\Desktop\AEBN20\Energy_simu\Initial_Results.xls')
