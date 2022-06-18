import xlrd
import xlwt

month_hour = [744, 672, 744, 720, 744, 720, 744, 744, 720, 744, 720, 744]

wb_simulation_use = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\elec_use_hourly.xls')
wb_initial_use = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\LCC.xls')
sheet_simulation_use = wb_simulation_use.sheet_by_name('electricity use hourly')

wb_PV_price = xlrd.open_workbook(r'C:\Users\Adam\Desktop\AEBN20\LCC Calc\xlwtxlrd\PV_prices.xls')
sheet_PV_production = wb_PV_price.sheet_by_name('Hourly profiles')

wb_PV_data_hour_monthly = xlwt.Workbook()
sheet_production = wb_PV_data_hour_monthly.add_sheet("PV production")
sheet_elec_load = wb_PV_data_hour_monthly.add_sheet("electricity load")


##check if production is more than purchase
PV_lar_prod_check = []
PV_lar_combi_check = []
PV_smal_prod_check = []
PV_smal_combi_check = []
PV_big_prod_check = []
PV_big_combi_check = []

# get prices for months
selling_price = []
buying_price = []
tax_compen = []
for ii in range(1, 13):
    sheet_price = wb_PV_price.sheet_by_name('Blad{}'.format(ii))
    selling_price.append(sheet_price.cell_value(6, 41))
    buying_price.append(sheet_price.cell_value(13, 41))
    tax_compen.append(sheet_price.cell_value(15,41))
# print(selling_price)
# print(buying_price)
# print(tax_compen)


# get annually production
prod_lar = []
prod_smal = []
prod_big = []
for zz in range(sheet_PV_production.nrows-2):
    prod_lar1 = sheet_PV_production.cell_value(zz+2,1)
    prod_smal1 = sheet_PV_production.cell_value(zz+2,3)
    prod_big1 = sheet_PV_production.cell_value(zz + 2, 5)
    prod_lar.append(prod_lar1)
    prod_smal.append(prod_smal1)
    prod_big.append(prod_big1)

# for 140 simulation
for mm in range(sheet_simulation_use.ncols-1):
    simucase = []
    for dd in range(sheet_simulation_use.nrows-1):

        # get annually simulation use
        simucase_1 = sheet_simulation_use.cell_value(dd+1, mm+1)
        simucase.append(simucase_1)


    # list for 12 months of elec saving/sold
    saved_elec_lar_yr = []
    saved_elec_smal_yr = []
    saved_elec_big_yr = []
    sold_elec_lar_yr = []
    sold_elec_smal_yr = []
    sold_elec_big_yr = []

    # separate annual data into monthly
    monthly_savings_kWh = []
    monthly_savings_SEK = []
    m = 0

    # sell/buy annually kWh amount
    sell_lar_kWh = []
    sell_smal_kWh = []
    sell_big_kWh = []
    buy_lar_kWh = []
    buy_smal_kWh = []
    buy_big_kWh = []
    for i in month_hour:
        m = m+i
        z = m-i
        # separate simulation data monthly
        monthly_use_simulation = simucase[z:m]

        # separate produce data monthly
        prod_lar_monthly = prod_lar[z:m]
        prod_smal_monthly = prod_smal[z:m]
        prod_big_monthly = prod_big[z:m]

        # write production\use data


        # Large system---calculate selling and buying electricity amount
        lar_sold_kWh = 0
        lar_buy_kWh = 0.
        lar_buy_amount = []
        for xx in range(len(prod_lar_monthly)):
            if prod_lar_monthly[xx] - monthly_use_simulation[xx] > 0:
                lar_sold_kWh = lar_sold_kWh + (prod_lar_monthly[xx] - monthly_use_simulation[xx])
            else:
                lar_buy_kWh = lar_buy_kWh + (-(prod_lar_monthly[xx] - monthly_use_simulation[xx]))

        # Small system---calculate selling and buying electricity amount
        smal_sold_kWh = 0
        smal_buy_kWh = 0
        smal_buy_amount = []
        for xx in range(len(prod_smal_monthly)):
            if prod_smal_monthly[xx] - monthly_use_simulation[xx] > 0:
                smal_sold_kWh = smal_sold_kWh + (prod_smal_monthly[xx] - monthly_use_simulation[xx])
            else:
                smal_buy_kWh = smal_buy_kWh + (-(prod_smal_monthly[xx] - monthly_use_simulation[xx]))

        # Big system---calculate selling and buying electricity amount
        big_sold_kWh = 0
        big_buy_kWh = 0
        big_buy_amount = []
        for xx in range(len(prod_big_monthly)):
            if prod_big_monthly[xx] - monthly_use_simulation[xx] > 0:
                big_sold_kWh = big_sold_kWh + (prod_big_monthly[xx] - monthly_use_simulation[xx])
            else:
                big_buy_kWh = big_buy_kWh + (-(prod_big_monthly[xx] - monthly_use_simulation[xx]))
        # lar_buy_amount.append(big_buy_kWh)
        # smal_buy_amount.append(smal_buy_kWh)
        # big_buy_amount.append(big_buy_kWh)



        # Saved/sold elec amount
        saved_elec_lar = sum(monthly_use_simulation)- lar_buy_kWh
        saved_elec_smal = sum(monthly_use_simulation) - smal_buy_kWh
        saved_elec_big = sum(monthly_use_simulation) - big_buy_kWh
        saved_elec_lar_yr.append(saved_elec_lar)
        saved_elec_smal_yr.append(saved_elec_smal)
        saved_elec_big_yr.append(saved_elec_big)
        sold_elec_lar_yr.append(lar_sold_kWh)
        sold_elec_smal_yr.append(smal_sold_kWh)
        sold_elec_big_yr.append(big_sold_kWh)


        sell_lar_kWh.append(lar_sold_kWh)
        sell_smal_kWh.append(smal_sold_kWh)
        sell_big_kWh.append(big_sold_kWh)
        buy_lar_kWh.append(lar_buy_kWh)
        buy_smal_kWh.append(smal_buy_kWh)
        buy_big_kWh.append(big_buy_kWh)


    # Calculate annual savings
    Annual_lar_saving_SEK = 0
    Annual_smal_saving_SEK = 0
    Annual_big_saving_SEK = 0
    for mo in range(len((sold_elec_smal_yr))):
        Annual_lar_saving_SEK = Annual_lar_saving_SEK + saved_elec_lar_yr[mo] * buying_price[mo] + sold_elec_lar_yr[mo] * selling_price[mo] + sold_elec_lar_yr[mo] * tax_compen[mo] * selling_price[mo]
        Annual_smal_saving_SEK = Annual_smal_saving_SEK + saved_elec_smal_yr[mo] * buying_price[mo] + sold_elec_smal_yr[mo] * selling_price[mo] +sold_elec_smal_yr[mo] * tax_compen[mo] * selling_price[mo]
        Annual_big_saving_SEK = Annual_big_saving_SEK + saved_elec_big_yr[mo] * buying_price[mo] + sold_elec_big_yr[mo] * selling_price[mo] +sold_elec_big_yr[mo] * tax_compen[mo] * selling_price[mo]

    ### check over production
    PV_lar_combi_check.append('{}_Large system_'.format(sheet_simulation_use.cell_value(0,mm+1)))
    if sum(sell_lar_kWh) > sum(buy_lar_kWh):
        PV_lar_prod_check.append('Over Production!')
        a = 'Over Production!'
    else:
        PV_lar_prod_check.append('NICE')
        a = 'NICE!'

    PV_smal_combi_check.append('{}_Small system_'.format(sheet_simulation_use.cell_value(0, mm + 1)))
    if sum(sell_smal_kWh) > sum(buy_smal_kWh):
        PV_smal_prod_check.append('Over Production!')
        b = 'Over Production!'
    else:
        PV_smal_prod_check.append('NICE')
        b = 'NICE!'

    PV_big_combi_check.append('{}_Big system_'.format(sheet_simulation_use.cell_value(0, mm + 1)))
    if sum(sell_big_kWh) > sum(buy_big_kWh):
        PV_big_prod_check.append('Over Production!')
        c = 'Over Production!'
    else:
        PV_big_prod_check.append('NICE')
        c = "NICE!"


    # ## Check over production
    # print(PV_lar_combi_check[mm]+a)
    # print(PV_smal_combi_check[mm] + b)
    # print(PV_big_combi_check[mm]+ c)

    # print(sell_lar_kWh)
    # print(buy_lar_kWh)
    # print(lar_sold_kWh)

# print(PV_lar_combi_check)
# print(PV_lar_prod_check)
# print(PV_smal_prod_check)
# print(PV_big_prod_check)


    # print(lar_sold_kWh)
    # print(lar_buy_kWh)


        # if mo == 0:
        #     break

    # # print SEK savings
    # print('{}_Large system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(Annual_lar_saving_SEK))
    # print('{}_Small system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(Annual_smal_saving_SEK))
    # print('{}_Big system_'.format(sheet_simulation_use.cell_value(0, mm + 1)) + str(Annual_big_saving_SEK))

    # print(Annual_lar_saving_SEK)
    # print(Annual_smal_saving_SEK)
    # print(Annual_big_saving_SEK)

    # # print kWh results
    # print('{}_Large system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(sum(saved_elec_lar_yr))+'_'+str(sum(sold_elec_lar_yr)))
    # print('{}_Small system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(sum(saved_elec_smal_yr))+'_'+str(sum(sold_elec_smal_yr)))
    # print('{}_Big system_'.format(sheet_simulation_use.cell_value(0, mm + 1)) + str(sum(saved_elec_big_yr)) + '_' + str(sum(sold_elec_big_yr)))

    # # print purchased electricity
    # print('{}_Large system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(sum(buy_lar_kWh)))
    # print('{}_Small system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(sum(buy_smal_kWh)))
    # print('{}_Big system_'.format(sheet_simulation_use.cell_value(0,mm+1)) + str(sum(buy_big_kWh)))


    # if mm == 0:
    #     break