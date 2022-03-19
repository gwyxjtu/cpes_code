import os
import os.path

import numpy as np
import xlrd
import xlwt as xlwt

from location_to_province import location_to_province

rootdir = 'E:/git_eppy_code/git_eppy_code/res1124'#文件夹目录
aimdir = 'E:/论文/生成负载/me'#文件夹目录
province = []
city = []
apartmentfilenamelist = ['']*500
hotelfilenamelist = ['']*500
officefilenamelist = ['']*500
restaurantfilenamelist = ['']*500
def data_remode(rootdir):
    for parent, dirnames, filenames in os.walk(rootdir):
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


if __name__ == '__main__':
    aimfilename = 'egqr.xls'
    province,city,restaurantfilenamelist,officefilenamelist,hotelfilenamelist,apartmentfilenamelist = data_remode(rootdir)
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
