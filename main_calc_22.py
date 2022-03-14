
import gurobipy as gp
from gurobipy import GRB
import numpy as np
import xlwt
import xlrd
import random
import time
import csv
from chicken_plan_zxxc_plan_h2devices import *
#chicken_plan_zxxc_plan_h2devices
from chicken_op import *
from load_generation import *
import json
import os
import pprint






def save_json(j,name):
    jj = json.dumps(j)
    f = open(res_dict+name+".json",'w')
    f.write(jj)
    f.close()
    return 0
if __name__ == '__main__':
    tem_env = 0#环境温度，后续补上
    #print(m_date)#main_input_zxxc_plan_h2devices
    with open("main_input_zxxc_plan_h2devices.json",encoding = "utf-8") as load_file:
        input_json = json.load(load_file)

    #dict_load = get_load()
    dict_load = get_load_new(input_json["load"])
    #dict_load = get_load(input_json["load"])

    #exit(0)
    #买电，卖电，买氢

    res1,grid_planning_output_json,grid_operation_output_json_plan,device_cap1 = planning_problem(dict_load, [1,input_json["load"]['power_sale_state']['grid'],input_json["load"]['hydrogen_state']['grid']], input_json)
    # device_cap1 = {'area_pv': 2052.0416969482167,
    #      'area_sc': 0.0,
    #      'hst': 29.738068123876026,
    #      'm_ct': 313651.26633118524,
    #      'm_ht': 10231.47838154615,
    #      'num_gtw': 1.0,
    #      'p_co': 5.20044466316282,
    #      'p_eb_max': 0.0,
    #      'p_el_max': 161.6199354558481,
    #      'p_fc_max': 267.5702413120801,
    #      'p_hp_max': 106.85655111531212,
    #      'p_hpg_max': 7.0}
    pprint.pprint(device_cap1)
    print(grid_planning_output_json['equipment_cost'],grid_planning_output_json['receive_year'])
    #grid_operation_output_json,flag = operating_problem(dict_load, device_cap1,[1,1,1],tem_env,input_json,8760)
    flag = 1
    if flag == 1:
        print("grid_g")
        grid_operation_output_json = grid_operation_output_json_plan


    #res2,itgrid_planning_output_json,isloate_operation_output_json_plan,device_cap2 = planning_problem(dict_load, [0,input_json["load"]['power_sale_state']['grid'],input_json["load"]['hydrogen_state']['isloate']], input_json)
    #pprint.pprint(device_cap2)
    #
    #itgrid_operation_output_json,flag = operating_problem(dict_load, device_cap2,[0,1,1],tem_env,input_json,8760)
    # flag = 1
    # if flag == 1:
    #     print("isloate_g")
    #     itgrid_operation_output_json = isloate_operation_output_json_plan
    #print(111)
    print(grid_planning_output_json['equipment_cost'],grid_planning_output_json['receive_year'])
    #print(itgrid_planning_output_json['equipment_cost'],itgrid_planning_output_json['receive_year'])
    pprint.pprint(device_cap1)
    #pprint.pprint(device_cap2)
    pprint.pprint(grid_operation_output_json)
    pprint.pprint(grid_operation_output_json_plan)
    #pprint.pprint(itgrid_operation_output_json)
    #pprint.pprint(isloate_operation_output_json_plan)

    #output_json = operating_problem(dict_load, device_cap, 1, tmp_env, input_json)

    #output_json = operating_problem(dict_load, device_cap, 0, tmp_env, input_json)

    


    save_json(grid_planning_output_json,"grid_planning_output_json")
    save_json(grid_operation_output_json,"grid_operation_output_json")
    #save_json(itgrid_planning_output_json,"itgrid_planning_output_json")
    #save_json(itgrid_operation_output_json,"itgrid_operation_output_json")
    to_csv(res1,'test1' + '.xls')
    #to_csv(res2,'test2' + '.xls')

#[0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49,0.49]