import numpy as np
import xlwt
import xlrd
import random
import time
import csv
#from chicken_plan_zxxc_plan_h2devices import *
#chicken_plan_zxxc_plan_h2devices
from chicken_op import *
import json
import os
import pprint
res_dict = "doc/"
m_date = [31,28,31,30,31,30,31,31,30,31,30,31]
m_date = [sum(m_date[:i])*24 for i in range(12)]
m_date.append(8760)
gb = {
    "Apartment":{
    #大面积旅馆
        1:87,
        2:68,
        3:70,
        4:94,
        5:60
    },
    "Hotel":{
        1:87,
        2:75,
        3:78,
        4:95,
        5:55
    },
    "Office":{
        1:59,
        2:39,
        3:36,
        4:34,
        5:25
    },
    "restaurant":{
        1:87,
        2:75,
        3:78,
        4:95,
        5:55
    }
}
def read_load_file(filename,gb,load_sort,heat_mounth,cool_mounth):

    ele_load = [1 for i in range(8760)]
    g_demand = [0 for i in range(8760)]
    q_demand = [0 for i in range(8760)]
    r_solar =  [1 for i in range(8760+24)]
    for h in heat_mounth:
        #print(m_date[h-1],m_date[h])
        g_demand[m_date[h-1]+15*24:min(m_date[h]+15*24,8760)] = [1 for _ in range(m_date[h]-m_date[h-1])]
    if 12 in heat_mounth:
        g_demand[:15*24] = [1 for _ in range(15*24)]
        g_demand = g_demand[:8760]
    for cc in cool_mounth:
        q_demand[m_date[cc-1]:m_date[cc]] = [1 for _ in range(m_date[cc]-m_date[cc-1])]
    with open("load/"+filename) as officecsv:

        office = csv.DictReader(officecsv)
        i=0
        for row in office:
            ele_load[i] *= float(row['Electricity Load [J]'])
            q_demand[i] *= float(row['Cooling Load [J]'])
            g_demand[i] *= float(row['Heating Load [J]'])
            
            i+=1

    s = gb[load_sort]
    tmp_sum = sum(ele_load)
    kkk = s/tmp_sum
    #GB

    # g_demand = [g_demand[i]*kkk for i in range(8760)]
    # q_demand = [q_demand[i]*kkk for i in range(8760)]
    # ele_load = [ele_load[i]*kkk for i in range(8760)]
    print(sum(g_demand),sum(q_demand),sum(ele_load))
    print(len(g_demand),len(q_demand),len(ele_load))
    print(m_date)
    #print(g_demand[2870:2890])
    #exit(0)
    return ele_load,g_demand,q_demand




def get_load_new(load_dict):
    jing = float(load_dict["location"][0])
    wei = float(load_dict["location"][1])
    load_dict["load_sort"] = 5 if jing>106 and wei<25 else 2
    if jing <106:
        load_dict["load_sort"] = 4
    if wei>35:
        load_dict["load_sort"] = 3
    if wei >=40 or (jing<101 and wei>28):
        load_dict["load_sort"] = 1
    print(load_dict["load_sort"])
    load_apartment = "ApartmentMidRise.csv"
    load_hotel     = "HotelSmall.csv"
    load_office    = "OfficeMedium.csv"
    load_restaurant= "RestaurantSitDown.csv"
    if load_dict["load_sort"] == 1:
        load_apartment  = "Heilongjiang_Harbin_"+load_apartment  
        load_hotel      = "Heilongjiang_Harbin_"+load_hotel      
        load_office     = "Heilongjiang_Harbin_"+load_office     
        load_restaurant = "Heilongjiang_Harbin_"+load_restaurant 
    if load_dict["load_sort"] == 2:
        load_apartment  = "Hebei_Shijiazhuang_"+load_apartment  
        load_hotel      = "Hebei_Shijiazhuang_"+load_hotel      
        load_office     = "Hebei_Shijiazhuang_"+load_office     
        load_restaurant = "Hebei_Shijiazhuang_"+load_restaurant 

    if load_dict["load_sort"] == 3:
        load_apartment  = "Jiangsu_Nanjing_"+load_apartment  
        load_hotel      = "Jiangsu_Nanjing_"+load_hotel      
        load_office     = "Jiangsu_Nanjing_"+load_office     
        load_restaurant = "Jiangsu_Nanjing_"+load_restaurant 
    if load_dict["load_sort"] == 4:
        load_apartment  = "Hainan_Haikou_"+load_apartment  
        load_hotel      = "Hainan_Haikou_"+load_hotel      
        load_office     = "Hainan_Haikou_"+load_office     
        load_restaurant = "Hainan_Haikou_"+load_restaurant 
    if load_dict["load_sort"] == 5:
        load_apartment  = "Yunnan_Kunming_"+load_apartment  
        load_hotel      = "Yunnan_Kunming_"+load_hotel      
        load_office     = "Yunnan_Kunming_"+load_office     
        load_restaurant = "Yunnan_Kunming_"+load_restaurant 
    # print(load_apartment )
    # print(load_hotel     )
    # print(load_office    )
    # print(load_restaurant)
    # exit(0)
    
    #国标修正

    e1,g1,q1 = read_load_file(load_apartment,gb["Apartment"],load_dict["load_sort"],load_dict["heat_mounth"],load_dict["cold_mounth"])
    e2,g2,q2 = read_load_file(load_hotel,gb["Hotel"],load_dict["load_sort"],load_dict["heat_mounth"],load_dict["cold_mounth"])
    e3,g3,q3 = read_load_file(load_office,gb["Office"],load_dict["load_sort"],load_dict["heat_mounth"],load_dict["cold_mounth"])
    e4,g4,q4 = read_load_file(load_restaurant,gb["restaurant"],load_dict["load_sort"],load_dict["heat_mounth"],load_dict["cold_mounth"])
    #print(load_dict["load_sort"]["building_area"])
    sum_rate = load_dict["building_area"]["apartment"]+load_dict["building_area"]["hotel"]+load_dict["building_area"]["office"]+load_dict["building_area"]["restaurant"]
    rate1 = load_dict["load_area"]*load_dict["building_area"]["apartment"]/sum_rate
    rate2 = load_dict["load_area"]*load_dict["building_area"]["hotel"]/sum_rate
    rate3 = load_dict["load_area"]*load_dict["building_area"]["office"]/sum_rate
    rate4 = load_dict["load_area"]*load_dict["building_area"]["restaurant"]/sum_rate
    ele_load =  [e1[i]*rate1 +e2[i]*rate2 +e3[i]*rate3 +e4[i]*rate4 for i in range(len(e1))]
    g_demand =  [g1[i]*rate1 +g2[i]*rate2 +g3[i]*rate3 +g4[i]*rate4 for i in range(len(g1))]
    q_demand =  [q1[i]*rate1 +q2[i]*rate2 +q3[i]*rate3 +q4[i]*rate4 for i in range(len(q1))]
    #q_demand[:92*24] = [0 for i in range(92*24)]
    #q_demand[-61*24:] = [0 for i in range(61*24)]
    #g_demand[92*24:92*24+212*24] = [0 for i in range(212*24)]
    print("-----")
    #print(rate1)
    print(len(g1),len(q1),len(e1))
    print(sum(g1),sum(q1),sum(e1))
    print(sum(g_demand),sum(q_demand),sum(ele_load))
    print("-----")
    period = len(e1)
    min_err = 10000
    files=os.listdir(r'solar')

    for file in files:
        #for i in range(len(df)):
        lon=file.split('_')[2]
        lat=file.split('_')[3]
        # print(((float(lat)-float(df.iloc[:,2][i]))**2+(float(lon)-float(df.iloc[:,3][i]))**2))
        #print(load_dict["location"])
        err = (float(lat)-float(load_dict["location"][0]))**2+(float(lon)-float(load_dict["location"][1]))**2
        if err<min_err:
            min_err = err
            final_file = file
        #print(float(lat),float(lon))
        #print(err,min_err)
    #print(final_file,min_err)
    r_solar =  [0 for i in range(8760+24)]
    with open("solar/"+final_file) as renewcsv:
        renewcsv.readline()
        renewcsv.readline()
        renewcsv.readline()
        renew = csv.DictReader(renewcsv)
        
        i=0
        for row in renew:

            r_solar[i] += float(row['electricity'])
    
            i+=1
    r_solar = r_solar[-8:]+r_solar[:-8]
    print("ori")
    print(sum(g_demand),sum(q_demand),sum(ele_load))
    print(max(g_demand),max(q_demand),max(ele_load))
    if load_dict["yearly_power"] !=0:
        #print(1)
        if load_dict["ele_type"] == 0:
            tmp_sum = sum(ele_load)+sum(q_demand)/4+sum(g_demand)/0.95
            kkk = load_dict["yearly_power"] /tmp_sum
            g_demand = [g_demand[i]*kkk for i in range(period)]
            q_demand = [q_demand[i]*kkk for i in range(period)]
            ele_load = [ele_load[i]*kkk for i in range(period)]
        else:
            kkk = (load_dict["yearly_power"]-sum(q_demand)/4-sum(g_demand)/0.95) /sum(ele_load)

            ele_load = [ele_load[i]*kkk for i in range(period)]
    if load_dict["power_peak"]['flag'] == 1:
        tmp_sum = max(g_demand)+0.000001
        kkk = load_dict["power_peak"]['g'] / tmp_sum
        g_demand = [g_demand[i]*max(kkk-load_dict["power_peak"]['flag_shear_g']*(tmp_sum-g_demand[i])/(g_demand[i]/load_dict["power_peak"]['shear']+0.01),1) for i in range(period)]
        #g_demand = [g_demand[i]*kkk for i in range(period)]
        

        tmp_sum = max(ele_load)+0.000001
        kkk = load_dict["power_peak"]['ele'] / tmp_sum
        ele_load = [ele_load[i]*max(kkk-load_dict["power_peak"]['flag_shear_e']*(tmp_sum-ele_load[i])/(ele_load[i]/load_dict["power_peak"]['shear']+0.01),1)  for i in range(period)]
        #ele_load = [ele_load[i]*kkk for i in range(period)]

        tmp_sum = max(q_demand) +0.000001
        kkk = load_dict["power_peak"]['q'] / tmp_sum
        q_demand = [q_demand[i]*max(kkk-load_dict["power_peak"]['flag_shear_q']*(tmp_sum-q_demand[i])/(q_demand[i]/load_dict["power_peak"]['shear']+0.01),1) for i in range(period)]
        #q_demand = [q_demand[i]*kkk for i in range(period)]
    print("new")
    print(max(g_demand),max(q_demand),max(ele_load))
    print(sum(g_demand),sum(q_demand),sum(ele_load))
    #exit(0)
    dict_load = {'ele_load': ele_load, 'g_demand': g_demand, 'q_demand': q_demand, 'r_solar': r_solar, 'load_sort':load_dict["load_sort"]}
    #exit(0)
    to_csv(dict_load,"test_load.csv")
    return dict_load

def get_load(load_dict):
    min_err = 10000
    files=os.listdir(r'solar')
    period = 8760

    #book_spr = xlrd.open_workbook('cspringdata.xlsx')
    #book_sum = xlrd.open_workbook('csummerdata.xlsx')
    #book_aut = xlrd.open_workbook('cautumndata.xlsx')
    #book_win = xlrd.open_workbook('cwinterdata.xlsx')
    for file in files:
        #for i in range(len(df)):
        lon=file.split('_')[2]
        lat=file.split('_')[3]
        # print(((float(lat)-float(df.iloc[:,2][i]))**2+(float(lon)-float(df.iloc[:,3][i]))**2))
        #print(load_dict["location"])
        err = (float(lat)-float(load_dict["location"][0]))**2+(float(lon)-float(load_dict["location"][1]))**2
        if err<min_err:
            min_err = err
            final_file = file
        #print(float(lat),float(lon))
        #print(err,min_err)
    #print(final_file,min_err)
    r_solar =  [0 for i in range(8760+24)]
    with open("solar/"+final_file) as renewcsv:
        renewcsv.readline()
        renewcsv.readline()
        renewcsv.readline()
        renew = csv.DictReader(renewcsv)
        
        i=0
        for row in renew:

            r_solar[i] += float(row['electricity'])
    
            i+=1
    r_solar = r_solar[-8:]+r_solar[:-8]

    #print(len(r_solar),len(t_env_indoor),len(t_env_outdoor))
    #print(max(r_solar))
    #exit(0)
    # book = xlrd.open_workbook('new_xls.xlsx')
    # data = book.sheet_by_index(0)
    # for l in range(1,8761):
    #     q_demand.append(data.cell(l,1).value)
    #     g_demand.append(data.cell(l,2).value)
    #     m_demand.append(data.cell(l,3).value)
    #     ele_load.append(data.cell(l,4).value)
    # q_demand = [0 if num == '' else num for num in q_demand]
    # g_demand = [0 if num == '' else num for num in g_demand]
    # m_demand = [0 if num == '' else num for num in m_demand]
    # ele_load = [0 if num == '' else num for num in ele_load]
    #print(q_demand)
    #exit(0)
    # for l in range(1,8761):
    #     q_demand.append(float(data.cell(l,1).value))
    #     g_demand.append(float(data.cell(l,2).value))
    #     m_demand.append(float(data.cell(l,3).value)/(30*c))
    #     ele_load.append(float(data.cell(l,4).value))
    ele_load = [10 for i in range(8760)]
    # tmp_mon = [959.8661111,920.8792593,819.3777778,701.4557407,641.8785185,547.9914815,530.1938889,531.7501852,650.3240741,737.0631481,839.3162963,961.3816667]
    # tmp = []
    # tmp+=[[0]*16 + [tmp_mon[0]/2, tmp_mon[0]/2]+[tmp_mon[0]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[1]/2, tmp_mon[1]/2]+[tmp_mon[1]*0.025  ]*4+[0]*2 for _ in range(28)]
    # tmp+=[[0]*16 + [tmp_mon[2]/2, tmp_mon[2]/2]+[tmp_mon[2]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[3]/2, tmp_mon[3]/2]+[tmp_mon[3]*0.025  ]*4+[0]*2 for _ in range(30)]
    # tmp+=[[0]*16 + [tmp_mon[4]/2, tmp_mon[4]/2]+[tmp_mon[4]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[5]/2, tmp_mon[5]/2]+[tmp_mon[5]*0.025  ]*4+[0]*2 for _ in range(30)]
    # tmp+=[[0]*16 + [tmp_mon[6]/2, tmp_mon[6]/2]+[tmp_mon[6]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[7]/2, tmp_mon[7]/2]+[tmp_mon[7]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[8]/2, tmp_mon[8]/2]+[tmp_mon[8]*0.025  ]*4+[0]*2 for _ in range(30)]
    # tmp+=[[0]*16 + [tmp_mon[9]/2, tmp_mon[9]/2]+[tmp_mon[9]*0.025  ]*4+[0]*2 for _ in range(31)]
    # tmp+=[[0]*16 + [tmp_mon[10]/2,tmp_mon[10]/2]+[tmp_mon[10]*0.025]*4+[0]*2 for _ in range(30)]
    # tmp+=[[0]*16 + [tmp_mon[11]/2,tmp_mon[11]/2]+[tmp_mon[10]*0.025]*4+[0]*2 for _ in range(31)]
    
    #g_demand = [j for i in tmp for j in i]
    g_demand = [[0]*16 + [375/2, 375/2]+[375*0.05  ]*4+[0]*2 for _ in range(365)]
    g_demand = [j for i in g_demand for j in i]
    #print(g_demand)
    q_demand = [0 for i in range(8760)]
    #r_solar =  [0 for i in range(8760+24)]

    print(len(g_demand),len(q_demand),len(ele_load))
    print(sum(g_demand),sum(q_demand),sum(ele_load))
    # print(len(r_solar))
    # exit(0)

    dict_load = {'ele_load': ele_load, 'g_demand': g_demand, 'q_demand': q_demand, 'r_solar': r_solar,'load_sort':load_dict["load_sort"]}
    to_csv(dict_load,"yt.xls")
    return dict_load

def to_csv(res,filename):
    items = list(res.keys())
    wb = xlwt.Workbook()
    total = wb.add_sheet('garden')
    for i in range(len(items)):
        total.write(0,i,items[i])
        if type(res[items[i]]) == list:
            sum = 0
            print(items[i])
            for j in range(len(res[items[i]])):
                total.write(j+2,i,(res[items[i]])[j])
                # sum += (res[items[i]])[j]
            # total.write(1,i,sum)
        else:
            print(items[i])
            total.write(1,i,res[items[i]])

    #filename = 'res/chicken_plan_2_load_1' + '.xls'
    wb.save(res_dict+filename)
