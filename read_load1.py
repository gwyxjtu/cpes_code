import csv
import json
import pypinyin
import os
import os.path
import xlrd
import xlwt as xlwt
rootdir = os.path.dirname(__file__) #文件夹目录
rootdir = rootdir.replace('\\','/')
print(rootdir)
resrootdir = os.path.join(os.path.dirname(__file__) + '/res1124/')#文件夹目录
resrootdir = resrootdir.replace('\\','/')
print(resrootdir)

import requests
import json
def location_to_province(Longitude,latitude):
    key = 'GjG3XAdmywz7CyETWqHwIuEC6ZExY6QT'
    location = str(Longitude)+','+str(latitude)
    print(location)
    # r = requests.get(url='http://api.map.baidu.com/geocoder/v2/', params={'location':location, 'ak':key,'output':'json'})
    # result = r.json()
    url = 'http://api.map.baidu.com/geocoder/v2/'
    data = {
        'ak': key,
        'location': location,
        'output': 'json'
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36',
        'Referer': 'https://www.google.com/'
    }

    res = requests.get(url, headers=headers, params=data)  # 添加请求头
    print(res.text)
    try:
        json.loads(res.text)
        result = json.loads(res.text)
        print(result)
        province = result['result']['addressComponent']['province']
        city = result['result']['addressComponent']['city']
        print(province)
        return province, city
    except ValueError:
        province = 0
        city = 0
        return '0','0'
def fenqu(wei,jing):
    load_sort = 5 if jing > 106 and wei < 25 else 2
    if jing < 106:
        load_sort = 4
    if wei > 35:
        load_sort = 3
    if wei >= 40 or (jing < 101 and wei > 28):
        load_sort = 1
    print (load_sort)
    return load_sort

def data_remode(resrootdir):
    province = []
    city = []
    apartmentfilenamelist = [''] * 500
    hotelfilenamelist = [''] * 500
    officefilenamelist = [''] * 500
    restaurantfilenamelist = [''] * 500
    for parent, dirnames, filenames in os.walk(resrootdir):
        for filename in filenames:
            data = filename.split('_')
            #ASHRAE9012016_RestaurantFastFood_Denver_CHN_Henan.Xinyang.572970_CSWD
            type = data[1]
            location = data[4]
            data1 = location.split('.')
            province0 = data1[0]
            city0 = data1[1]
            if city.count(city0) == 0:
                province.append(province0)
                city.append(city0)
            suoyin = city.index(city0)
            if type =="RestaurantFastFood":restaurantfilenamelist[suoyin]=filename
            elif type =="OfficeMedium":officefilenamelist[suoyin]=filename
            elif type =="HotelSmall":hotelfilenamelist[suoyin]=filename
            elif type =="ApartmentHighRise":apartmentfilenamelist[suoyin]=filename
    return province,city,restaurantfilenamelist,officefilenamelist,hotelfilenamelist,apartmentfilenamelist


def add_eqpr():
    aimfilename = 'egqr.xls'
    province,city,restaurantfilenamelist,officefilenamelist,hotelfilenamelist,apartmentfilenamelist = data_remode(resrootdir)
    wb = xlwt.Workbook()
    total = wb.add_sheet('egqr')
    for i in range(len(province)):
        total.write(i,0,province[i])
        total.write(i,1,city[i])
        total.write(i,2,restaurantfilenamelist[i])
        total.write(i,3,officefilenamelist[i])
        total.write(i,4,hotelfilenamelist[i])
        total.write(i,5,apartmentfilenamelist[i])
    wb.save(aimfilename)
    return 0




#黑龙江 [46.5969,125.1015]
#广东 [22.2737,113.5721]
#陕西 [34.2204,109.1115]
#北京 [39.9062,116.3913]
#安徽 [31.8228,117.2218]
#云南 [24.8843,102.8324]
theta ={
    "Heilongjiang":[2,0.7],
    "Guangdong":[1,1],
    "Shaanxi":[1.8,0.8],
    "Beijing":[1.8,0.8],
    "Anhui":[1.6,1],
    "Yunnan":[1.5,0.7]
}

m_date = [31,28,31,30,31,30,31,31,30,31,30,31]
m_date = [sum(m_date[:i])*24 for i in range(12)]
m_date.append(8760) #每个月第一个小时的索引
base_apartment = [40,60,150]#W/m2
base_hotel = [55,60,150]#W/m2
base_office = [50,60,180]#W/m2
base_restaurant = [60,60,350]#W/m2

def peakbasecorrectload(pinprovince,base_ele_load, base_g_demand, base_q_demand,building_area,load_area):
    s_all = float(building_area['apartment']) + float(building_area['hotel']) + float(building_area['office']) + float(
        building_area['restaurant'])
    s_apartment = float(building_area['apartment']) / s_all
    s_hotel = float(building_area['hotel']) / s_all
    s_office = float(building_area['office']) / s_all
    s_restaurant = float(building_area['restaurant']) / s_all
    base_e = base_apartment[0] * s_apartment + base_hotel[0] * s_hotel + base_office[0] * s_office + base_restaurant[0] * s_restaurant
    base_g = base_apartment[1] * s_apartment + base_hotel[1] * s_hotel + base_office[1] * s_office + base_restaurant[1] * s_restaurant
    base_q = base_apartment[2] * s_apartment + base_hotel[2] * s_hotel + base_office[2] * s_office + base_restaurant[2] * s_restaurant
    sum_e_cankao = base_e * load_area  / 1000
    sum_g_cankao = base_g * load_area  /1000 *theta[pinprovince][0]
    sum_q_cankao = base_q * load_area  /1000 *theta[pinprovince][1]
    base_ele_load, base_g_demand, base_q_demand = peakcorrectload(base_ele_load, base_g_demand, base_q_demand, sum_e_cankao, sum_g_cankao, sum_q_cankao)
    return base_ele_load, base_g_demand, base_q_demand






def gqmonthcorrectload(g_demand, q_demand,heat_mounth,cold_mounth):
    gn_demand = [0 for i in range(8760)]
    qn_demand = [0 for i in range(8760)]
    z_heat_mounth = [0 for i in range(8760)]
    z_cold_month = [0 for i in range(8760)]
    start_h_m = int(heat_mounth.split('于')[1].split('月')[0])
    start_h_d = int(heat_mounth.split('于')[1].split('月')[1].split('日')[0])
    end_h_m = int(heat_mounth.split('于')[2].split('月')[0])
    end_h_d = int(heat_mounth.split('于')[2].split('月')[1].split('日')[0])
    start_c_m = int(cold_mounth.split('于')[1].split('月')[0])
    start_c_d = int(cold_mounth.split('于')[1].split('月')[1].split('日')[0])
    end_c_m = int(cold_mounth.split('于')[2].split('月')[0])
    end_c_d = int(cold_mounth.split('于')[2].split('月')[1].split('日')[0])
    start_h_index = m_date[start_h_m-1] + 24 * (start_h_d-1)
    end_h_index = m_date[end_h_m-1] + 24 * (end_h_d-1)
    start_c_index = m_date[start_c_m-1] + 24 * (start_c_d - 1)
    end_c_index = m_date[end_c_m-1] + 24 * (end_c_d - 1)
    print(start_h_index,end_h_index,start_c_index,end_c_index)
    if end_h_index >= start_h_index:
        gn_demand[start_h_index:end_h_index] = g_demand[start_h_index:end_h_index]
        z_heat_mounth[start_h_index:end_h_index] = [1 for i in range(end_h_index - start_h_index)]
    else:
        gn_demand[0:end_h_index] = g_demand[0:end_h_index]
        gn_demand[start_h_index:8760] = g_demand[start_h_index:8760]
        z_heat_mounth[0:end_h_index] = [1 for i in range(end_h_index)]
        z_heat_mounth[start_h_index:8760] = [1 for i in range(8760 - end_h_index)]
    if end_c_index >= start_c_index:
        qn_demand[start_c_index:end_c_index] = q_demand[start_c_index:end_c_index]
        z_cold_month[start_c_index:end_c_index] = [1 for i in range(end_c_index - start_c_index)]
    else:
        qn_demand[0:end_c_index] = q_demand[0:end_c_index]
        qn_demand[start_c_index:8760] = q_demand[start_c_index:8760]
        z_cold_month[0:end_c_index] = [1 for i in range(end_c_index)]
        z_cold_month[start_c_index:8760] = [1 for i in range(8760 - end_c_index)]
    return gn_demand, qn_demand,  z_heat_mounth, z_cold_month



def peakcorrectload(base_ele_load, base_g_demand, base_q_demand,peak_ele,peak_g,peak_q):
    ele_max , max_g , max_q = max(base_ele_load) , max(base_g_demand), max(base_q_demand)
    ele_load = [0 for i in range(8760)]
    g_demand = [0 for i in range(8760)]
    q_demand = [0 for i in range(8760)]
    for i in range(0,8760):
        ele_load[i] = base_ele_load[i] * peak_ele / ele_max
        if max_g == 0:
            g_demand[i] = 0
        else:
            g_demand[i] = base_g_demand[i] * peak_g / max_g
        if max_q == 0:
            q_demand[i] = 0
        else:
            q_demand[i] = base_q_demand[i] * peak_q / max_q
    return ele_load, g_demand, q_demand

def sumcorrectload(base_ele_load, base_g_demand, base_q_demand,sum_ele,sum_g,sum_q):
    all_ele , all_g , all_q = sum(base_ele_load) , sum(base_g_demand), sum(base_q_demand)
    ele_load = [0 for i in range(8760)]
    g_demand = [0 for i in range(8760)]
    q_demand = [0 for i in range(8760)]
    for i in range(0,8760):
        ele_load[i] = base_ele_load[i] * sum_ele / all_ele
        g_demand[i] = base_g_demand[i] * sum_g / all_g
        q_demand[i] = base_q_demand[i] * sum_q / all_q
    print(sum(ele_load),sum(g_demand),sum(q_demand))
    return ele_load, g_demand, q_demand


def baseautoload(pinprovince,pincity,building_area,load_area):
    ele_load = [0 for i in range(8760)]
    g_demand = [0 for i in range(8760)]
    q_demand = [0 for i in range(8760)]
    r_ele_load,o_ele_load,h_ele_load,a_ele_load= [0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)]
    r_g_demand,o_g_demand,h_g_demand,a_g_demand= [0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)]
    r_q_demand,o_q_demand,h_q_demand,a_q_demand= [0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)]
    r_solar = [0 for i in range(8760)]
    r_r_solar,o_r_solar,h_r_solar,a_r_solar= [0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)],[0 for i in range(8760)]
    s_all = float(building_area['apartment'])+float(building_area['hotel'])+float(building_area['office'])+float(building_area['restaurant'])
    s_apartment = float(building_area['apartment'])/s_all
    s_hotel = float(building_area['hotel'])/s_all
    s_office = float(building_area['office'])/s_all
    s_restaurant = float(building_area['restaurant'])/s_all
    restaurantfilename,officefilename,hotelfilename,apartmentfilename =findloadname(pinprovince,pincity)
    with open(resrootdir+ restaurantfilename) as restaurantcsv:
        restaurant = csv.DictReader(restaurantcsv)
        i = 0
        for row in restaurant:
            r_ele_load[i] = float(row["Electricity Load [kwh]"])
            r_g_demand[i] = float(row["Heating Load [kwh]"])
            r_q_demand[i] = float(row["Cooling Load [kwh]"])
            r_r_solar[i] = float(row["Environment:Site Direct Solar Radiation Rate per Area [W/m2](Hourly)"])
            i += 1
    with open(resrootdir + officefilename) as officecsv:
        office = csv.DictReader(officecsv)
        i = 0
        for row1 in office:
            o_ele_load[i] = float(row1["Electricity Load [kwh]"])
            o_g_demand[i] = float(row1["Heating Load [kwh]"])
            o_q_demand[i] = float(row1["Cooling Load [kwh]"])
            o_r_solar[i] = float(row1["Environment:Site Direct Solar Radiation Rate per Area [W/m2](Hourly)"])
            i += 1
    with open(resrootdir + hotelfilename) as hotelcsv:
        hotel = csv.DictReader(hotelcsv)
        i = 0
        for row2 in hotel:
            h_ele_load[i] = float(row2["Electricity Load [kwh]"])
            h_g_demand[i] = float(row2["Heating Load [kwh]"])
            h_q_demand[i] = float(row2["Cooling Load [kwh]"])
            h_r_solar[i] = float(row2["Environment:Site Direct Solar Radiation Rate per Area [W/m2](Hourly)"])
            i += 1
    with open(resrootdir + apartmentfilename) as apartmentcsv:
        apartment = csv.DictReader(apartmentcsv)
        i = 0
        for row3 in apartment:
            a_ele_load[i] = float(row3["Electricity Load [kwh]"])
            a_g_demand[i] = float(row3["Heating Load [kwh]"])
            a_q_demand[i] = float(row3["Cooling Load [kwh]"])
            a_r_solar[i] = float(row3["Environment:Site Direct Solar Radiation Rate per Area [W/m2](Hourly)"])
            i += 1
    for i in range(0,8760):
        ele_load[i] = (a_ele_load[i] * s_apartment + h_ele_load[i] * s_hotel + o_ele_load[i] * s_office + r_ele_load[i] * s_restaurant)*load_area
        g_demand[i] = (a_g_demand[i] * s_apartment + h_g_demand[i] * s_hotel + o_g_demand[i] * s_office + r_g_demand[i] * s_restaurant)*load_area
        q_demand[i] = (a_q_demand[i] * s_apartment + h_q_demand[i] * s_hotel + o_q_demand[i] * s_office + r_q_demand[i] * s_restaurant)*load_area
        r_solar[i] = a_r_solar[i] * s_apartment + h_r_solar[i] * s_hotel + o_r_solar[i] * s_office + r_r_solar[i] * s_restaurant

    return ele_load, g_demand, q_demand, r_solar



def findloadname(pinprovince,pincity):
    book = xlrd.open_workbook('egqr.xls')
    data = book.sheet_by_index(0)
    sheng = data.col_values(0, start_rowx=0, end_rowx=None)
    shi = data.col_values(1, start_rowx=0, end_rowx=None)
    suoyin = sheng.index(pinprovince)
    restaurantfilename = data.cell_value(suoyin,2)
    officefilename = data.cell_value(suoyin,3)
    hotelfilename = data.cell_value(suoyin,4)
    apartmentfilename = data.cell_value(suoyin,5)
    return restaurantfilename,officefilename,hotelfilename,apartmentfilename


def pinyin(word):
    s = ''
    for i in pypinyin.pinyin(word, style=pypinyin.NORMAL):
        s += ''.join(i)
    return s

def file_to_list(fileaddress):
    ele_load = []
    g_demand = []
    q_demand = []
    r_solar = []
    book = xlrd.open_workbook(fileaddress)
    data = book.sheet_by_index(0)
    for l in range(1,8761):
        ele_load.append(data.cell(l,0).value)
        g_demand.append(data.cell(l,1).value)
        q_demand.append(data.cell(l,2).value)
        r_solar.append(data.cell(l,3).value)
    return ele_load, g_demand, q_demand, r_solar

def all_load(dict_load):
    add_eqpr()
    location = dict_load['location']
    load_sort = fenqu(location[0], location[1])
    #如果不自动生成文档，直接从路径读取
    if dict_load["autoload"]==0:
        ele_load, g_demand, q_demand, r_solar = file_to_list(dict_load["fileaddress"])
        aimfilename = 'addressload_result.xls'
        wb = xlwt.Workbook()
        total = wb.add_sheet('egqr')
        for i in range(8760):
            total.write(i + 1, 0, ele_load[i])
            total.write(i + 1, 1, g_demand[i])
            total.write(i + 1, 2, q_demand[i])
            total.write(i + 1, 3, r_solar[i])
        wb.save(aimfilename)
    # 如果自动生成文档，读参数开始输入
    else:
        #如果有给定的城市名称
        if dict_load['city']!="":
            province, city = dict_load['province'],dict_load['city']
        #如果只有经纬度
        else:
            location = dict_load['location']
            province, city = location_to_province(location[0], location[1])
            print(province,city)
            i = 0
            while city == '0' and i <20:
                province, city = location_to_province(location[0], location[1])
                i=i+1
            if city == '0' and i >19:
                return "网络异常，请修改配置文件增加city"
            else:
                pinprovince = pinyin(province)[0:-5]
                pincity = pinyin(city)[0:-3]
                pinprovince = pinprovince.capitalize()
                pincity = pincity.capitalize()
                if province == "陕西省":
                    pinprovince = "Shaanxi"
                elif province == "北京市":
                    pinprovince = "Beijing"
                elif province == "重庆市":
                    pinprovince = "Chongqing"
                elif province == "天津市":
                    pinprovince = "Tianjin"
                elif province == "西藏自治区":
                    pinprovince = "Tibet"
                elif province == "广西壮族自治区":
                    pinprovince = "Guangxi"
                elif province == "新疆维吾尔自治区":
                    pinprovince = "Xinjiang"
                elif province == "宁夏回族自治区":
                    pinprovince = "Ningxia"
                elif province == "内蒙古自治区":
                    pinprovince = "Nei.Mongol"
                elif province == "新疆维吾尔自治区":
                    pinprovince = "Xinjiang"
                elif province == "上海市":
                    pinprovince = "Shanghai"
                building_area = dict_load['building_area']
                load_area = dict_load['load_area']
                print(pinprovince)
                print(pincity)
                #基础负荷
                base_ele_load, base_g_demand, base_q_demand, r_solar = baseautoload(pinprovince,pincity,building_area,load_area)
                power_peak = dict_load['power_peak']
                peak_flag = power_peak['flag']
                peak_ele = power_peak['ele']
                peak_g = power_peak['g']
                peak_q = power_peak['q']
                power_sum = dict_load['power_sum']
                sum_flag = power_sum['flag']
                sum_ele = power_sum['ele']
                sum_g = power_sum['g']
                sum_q = power_sum['q']
                # 供热、供冷矫正
                heat_mounth = dict_load['heat_mounth']
                cold_mounth = dict_load['cold_mounth']
                base_g_demand, base_q_demand, z_heat_mounth, z_cold_month = gqmonthcorrectload(base_g_demand, base_q_demand, heat_mounth, cold_mounth)
                base_ele_load, base_g_demand, base_q_demand = peakbasecorrectload(pinprovince,base_ele_load, base_g_demand, base_q_demand,building_area,load_area)
                #峰值矫正
                if peak_flag == 1 and sum_flag ==0:
                    ele_load, g_demand, q_demand = peakcorrectload(base_ele_load, base_g_demand, base_q_demand,peak_ele,peak_g,peak_q)
                #总量矫正
                elif peak_flag == 0 and sum_flag == 1:
                    ele_load, g_demand, q_demand = sumcorrectload(base_ele_load, base_g_demand, base_q_demand,sum_ele,sum_g,sum_q)
                #不矫正
                elif peak_flag == 0 and sum_flag == 0:
                    ele_load, g_demand, q_demand = base_ele_load, base_g_demand, base_q_demand
                #有峰值有总量，优先峰值
                else:
                    ele_load, g_demand, q_demand = peakcorrectload(base_ele_load, base_g_demand, base_q_demand,peak_ele,peak_g,peak_q)
                return_load = {'ele_load': ele_load, 'g_demand': g_demand, 'q_demand': q_demand, 'r_solar': r_solar, 'load_sort':load_sort,'z_cold_mounth':z_cold_month,'z_heat_mounth':z_heat_mounth}
                return return_load
                #return ele_load, g_demand, q_demand, r_solar, z_heat_mounth, z_cold_month, load_sort

if __name__ == '__main__':
    # 读输入文件
    with open("main_input.json", encoding="utf-8") as load_file:
        input_json = json.load(load_file)
    dict_load = input_json["load"]
    ele_load, g_demand, q_demand, r_solar, z_heat_mounth, z_cold_month, load_sort = all_load(dict_load)
    aimfilename = 'autoload_result.xls'
    wb = xlwt.Workbook()
    total = wb.add_sheet('egqr')
    for i in range(8760):
        total.write(i+1,0,ele_load[i])
        total.write(i+1,1,g_demand[i])
        total.write(i+1,2,q_demand[i])
        total.write(i+1,3,r_solar[i])
        total.write(i+1,4,z_heat_mounth[i])
        total.write(i+1,5,z_cold_month[i])
    total.write(1, 6, load_sort)
    total.write(0, 0, "电负荷/kwh")
    total.write(0, 1, "热负荷/kwh")
    total.write(0, 2, "冷负荷/kwh")
    total.write(0, 3, "光照强度/W/m2")
    total.write(0, 4, "z_heat_month")
    total.write(0, 5, "z_cold_month")
    total.write(0, 6, "load_sort")

    wb.save(aimfilename)
