# cpes_code

本项目为半自动计算综合能源系统测算报告的测算部分代码。

# 使用方法

## 修改配置文件
通过`main_input.json`进行配置，主要是负荷和可再生能源和价格部分。

## 运行主计算程序
运行`main_calc.py`进行计算,计算结果被输出到doc子目录下。

## 修改文档相关文字
修改doc文件夹下的xlsx文件

## 运行文档生成程序
运行doc下的`main.py`输出`out.docx`，为最终输出的文档。


# 输入配置文件说明
配置文件包`main_input.json`。具体配置说明如下。


## 详细配置文件

### 负荷部分

输入,`building_area`为各个类型房屋的比例；`heat_mounth,cold_mounth`为供热和供冷的月份，`power_peak`为峰值电负荷，缺省为0;`yearly_power`为年用电总量;`load_area`为总用能面积，非常关键;`ele_type`如果为1，代表冷热负荷与总电量无关，只与面积有关，如果为0代表冷热负荷与总电量有关。  `location`为坐标经纬度，第一个为经度，第二个为纬度。用于获取当地光照水平以及供暖条件；`power_peak`是负荷峰值，flag1表示启用，0表示不启用；`shear`表示负荷削减程度越大说明总负荷越低。`hydrogen_state`表示是否允许买氢，0表示不允许，1表示允许。
```
    "load":{
        "building_area":{
            "apartment":1,
            "hotel":0,
            "office":0,
            "restaurant":0
        },
        "load_area":9936,
        "heat_mounth":[1,2,3,4,11,12],
        "cold_mounth":[],
        "hydrogen_state":{
            "grid":0,
            "isloate":0
        },
        "power_peak":{
            "flag":1,
            "ele":400,
            "g":800,
            "q":0,
            "shear":4
        },
        "yearly_power":0,
        "ele_type":0,
        "location":[92,31]
    },
```
### 计算模式部分
`grid`表示并网计算模式，`p_pur_state` 表示是否允许买电，`p_sol_state` 表示是否允许卖电，`h_pur_state` 表示是否允许买氢。 `isloate`表示离网运行模式，`flag`表示是否计算离网模式，h_pur_state 表示是否允许买氢。    `obj`为目标函数选项，`capex_sum`表示总投资成本，`capex_crf`代表年化投资 `opex`代表运行成本，如果为1代表这一项在目标函数中。
```
    "calc_mode":{
        "grid":{
            "p_pur_state":1,
            "p_sol_state":0,
            "h_pur_state":0
        },
        "isloate":{
            "flag":0,
            "h_pur_state":0
        },
        "obj":{
            "capex_sum":0,
            "capex_crf":1,
            "opex":1
        }
    },
```


### 可再生能源部分

`pv_existst`为已有光伏装机容量功率；`sc_existst`为已有的太阳能集热器面积；`pv_sc_max`为光伏板和集热器一共的最大面积。
```
    "renewable_energy":{
        "pv_existst":1000,
        "sc_existst":1000,
        "pv_sc_max":100000
    },
```

### 市场价格部分

第一项`TOU_power`为数组，代表24h分时电价，第二项`power_sale`为卖电价格；`gas_price`是买天然气价格；`heat_price`为供热价格，单位是元/平方米每月；`cold_price`为供冷价格；第五项`hydrogen_price`为氢价，单位是元/kg；`capex_max`是投资成本上限，第一项为并网运行的规划投资，第二项为离网运行的投资上限，不需要上限可以设置为非常大。
```
    "price":{
        "TOU_power":[0.6713,0.6713,0.6713,0.6713,0.6713,0.9935,0.9935,0.6713,0.6713,0.6713,0.3155,0.3155,0.3155,0.3155,0.3155,0.6713,0.6713,0.9935,0.9935,0.9935,0.9935,0.6713,0.6713,0.6713],
        "power_sale":0.9,
        //卖电价格
        "gas_price":1.2,
        "heat_price":12,
        "cold_price":16,


        "hydrogen_price":30,
        "capex_max":[400000000,10000000000]
    },
```

### 设备部分

设备包括氢能系统设备和水循环供热系统设备。

el代表电解槽设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。
运行参数中`beta_el`代表电解槽制氢效率，`500RtV`代表输出氢气质量压强换算系数，`U_max,U_min`代表输出压强上下限。`nm3_max`   `nm3_min`分别是电解槽规划容量的上限和下限，单位是标方每小时。
```
	"el"://电解槽
	    {
	        "power_max":500,
	        //最大规划容量
	        "power_min":0,
	        //最小规划容量
            "nm3_max":10000000,
            //最大标方
            "nm3_min":0,
            //最小标方数目

	        "cost":3360,
	        //设备单位投资价格，这里的单位是元/kw，换算成标方需要乘以50再除以11.2
	        "crf":7,
	        //设备寿命
            "pressure_upper": 3,
            "pressure_lower": 0.1,
	        //运行参数
	        "beta_el":0.03,
	        //制氢效率
	        "500RtV":0,
	        //输出氢气质量压强换算系数
	        "U_max":200,
	        "U_min":20,
	        //输出压强上下限
	    },
```

co代表压缩机设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。
运行参数中`beta_co`代表氢压机功耗。
```
    "co":
        {
            "power_max":5000,
            "power_min":0,
            "cost":1000,
            //投资单价，元/kw
            "crf":10,

            "beta_co":1.399
            //功耗效率
        },
```
hst代表储氢罐`sto_max`代表最大规划容量,`inout_max`代表单位时间最大进出氢气量,`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。`U_max,U_min`代表压强上下限。
```
	"hst"://储氢罐
	    {
	        "sto_max":5000,
	        //最大容量
	        "stp_min":0,
	        //最小规划容量
	        "cost":1791,
	        //设备单位投资价格，元/kw
	        "crf":10,
	        //设备寿命

            "pressure_upper": 35,
            //最大压强
            "pressure_lower": 0.5,
            //最小储氢压强

	        //运行参数
	        "inout_max":500,
	        //最大进出量
	        "U_max":350,
	        "U_min":5,
	        //储氢罐压强上下限
	    },
 
```

fc代表燃料电池设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。
运行参数中`eta_fc_p`代表产电效率，`eta_ex_g`代表制热效率，`theta_ex`代表换热效率。
```
    "fc":
        {
            "power_max":100000,
            "power_min":0,
            "cost":8000,
            //投资单价，元/kw
            "crf":10,
            "nominal_lower": 0,

            "eta_fc_p":15,
            "eta_ex_g":16.6,
            "theta_ex":0.95,

            "V_c": 10,
            "t_fc": 85
        },
```

sc代表太阳能集热器，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。运行参数中`beta_sc`代表效率，`miu_loss`代表损失率，`theta_ex`代表换热效率。

```
    "sc"://太阳能集热器
        {
            "area_max":10000,
            //最大规划容量
            "area_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格，元/m2
            "crf":20,
            //设备寿命

            //运行参数
            "beta_sc":0.55,
            //集热器效率
            "theta_ex":0.9,
            //换热效率
            "miu_loss":0.9,
            //损失率
        },

```
ht代表热水罐，ct代表冷水罐，`water_max`和`water_min`代表最大最小规划容量`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。
运行参数中`t_max`，`t_min`代表最大最小水罐温度，`t_supply`代表最小供热温度和最大供冷温度，`miu_loss`代表损失率，`theta_ex`代表换热效率。

```
    "ht"://热水罐
        {
            "water_max":10000,
            //最大规划容量
            "water_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格 元/kg
            "crf":20,
            //设备寿命

            //运行参数
            "t_max":82,
            //最高水温
            "t_min":4,
            //最低水温
            "t_supply":55,
            //最低供热温度
            "miu_loss":0.02,
            //损失率m2.K
        },
    "ct"://冷水罐
        {
            "water_max":10000,
            //最大规划容量
            "water_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格 元/kg
            "crf":20,
            //设备寿命

            //运行参数
            "t_max":21,
            //最高水温
            "t_min":4,
            //最低水温
            "t_supply":7,
            //最低水温
            "miu_loss":0.9,
            //损失率
        },
```

eb代表电热锅炉设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。`beta_eb`代表产热效率。

```
    "eb"://电热锅炉
        {
            "power_max":10000,
            //最大规划容量
            "power_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格 元/kw
            "crf":20,
            //设备寿命

            //运行参数
            "beta_eb":0.8,
            //产热效率
        },
```
hp代表空气源热泵设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。`beta_hpg`代表产热COP，`beta_hpq`代表制冷COP。

```
    "hp"://空气源热泵
        {
            "power_max":10000,
            //最大规划容量
            "power_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格 元/kw
            "crf":20,
            //设备寿命

            //运行参数
            "beta_hpg":3,
            //产热效率
            "beta_hpq":4,
            //制冷效率
        },
```

ghp代表地源热泵设备，`power_max`代表最大规划容量，`power_min`代表最小规划容量，`cost`代表设备单位投资价格，元/kwh，`crf`代表设备寿命。`beta_ghpg`代表产热COP，`beta_ghpq`代表制冷COP。
```
    "ghp"://地源热泵
        {
            "power_max":10000,
            //最大规划容量
            "power_min":0,
            //最小规划容量
            "cost":2000,
            //设备单位投资价格 元/kw
            "crf":20,
            //设备寿命

            //运行参数
            "beta_ghpg":3.4,
            //产热效率
            "beta_ghpq":3.4,
            //制冷效率
        },
```
hyd代表水电模块，flag代表是否使用，supply代表水电的电能是否可以直接用于能源站能源供应，peak表示水电峰值，-1表示按照默认数据，cost表示水电投资单价，crf表示水电寿命。
```
        "hyd":
            {
                "flag":0,
                "supply":0,
                "power_cost":0,

                "peak":-1,

                "cost":0,
                "crf":20
            },
```

gtw代表浅层地热井，前三个参数分别代表最大规划数目，单个地热井投资成本，地热井寿命，`beta_gtw`代表每个地热井的最大换热功率(kW)。
```
        "gtw":
            {
                "number_max":100000,
                "cost":30000,
                "crf":30,

                "beta_gtw":7
            },
```
### 碳排放参数

```
    "carbon":{
        //碳排放因子
        "alpha_h2":1.74,
        //氢气排放因子
        "alpha_e":0.5839,
        //电网排放因子kg/kwh
        "alpha_EO":0.8922,
        //减排项目基准排放因子
    }
```

## 计算程序输出

输出参数如下，输出到doc文件夹下
```
#并网规划输出
grid_planning_output_json = {
        ‘flag_isloate’:0,是否需要离网报告

        'ele_load_sum': 3000,  # 电负荷总量/kwh
        'g_demand_sum': 3000,  # 热负荷总量/kwh
        'q_demand_sum': 3000,  # 冷负荷总量/kwh
        'ele_load_max': 30,  # 电负荷峰值/kwh
        'g_demand_max': 30,  # 热负荷峰值/kwh
        'q_demand_max': 30,  # 冷负荷峰值/kwh
        'ele_load': ele_load,  # 电负荷8760h的分时数据/kwh 一个8760的数组
        'g_demand': g_demand,  # 热负荷8760h的分时数据/kwh 一个8760的数组
        'q_demand': q_demand,  # 冷负荷8760h的分时数据/kwh 一个8760的数组
        'r_solar': r_solar,  # 光照强度8760h的分时数据/kwh 一个8760的数组

        'num_gtw': 5,  # 地热井数目/个
        'p_fc_max': 300,  # 燃料电池容量/kw
        'p_hpg_max': 300,  # 地源热泵功率/kw
        'p_hp_max': 300,  # 空气源热泵功率/kw
        'p_eb_max': 300  # 电热锅炉功率/kw
        'p_el_max': 300,  # 电解槽功率/kw
        'hst': 100,  # 储氢罐容量/kg
        'm_ht': 100,  # 储热罐/kg
        'm_ct': 100,  # 冷水罐/kg
        'area_pv': 600,  # 光伏面积/m2
        'area_sc': 500,  # 集热器面积/m2
        'p_co': 300,  #氢压机功率/kw

        "equipment_cost": 1000,  #设备总投资/万元
        "receive_year": 10,  # 投资回报年限/年
}
#并网结合运行后的输出
grid_operation_output_json = {
        "operation_cost": 10,  # 年化运行成本/万元
        "cost_save_rate": 20,  #运行成本节约比例%
        "co2":5,  #总碳排/t
        "cer":20,  #碳减排量
        "cer_rate": 70,  # 碳减排率%
        "cer_perm2":2  #每平米的碳减排量/t
}

#离网规划输出
itgrid_planning_output_json = {
        'ele_load_sum': 3000,  # 电负荷总量/kwh
        'g_demand_sum': 3000,  # 热负荷总量/kwh
        'q_demand_sum': 3000,  # 冷负荷总量/kwh
        'ele_load_max': 30,  # 电负荷峰值/kwh
        'g_demand_max': 30,  # 热负荷峰值/kwh
        'q_demand_max': 30,  # 冷负荷峰值/kwh
        'ele_load': ele_load,  # 电负荷8760h的分时数据/kwh
        'g_demand': g_demand,  # 热负荷8760h的分时数据/kwh
        'q_demand': q_demand,  # 冷负荷8760h的分时数据/kwh
        'r_solar': r_solar,  # 光照强度8760h的分时数据/kwh

        'num_gtw': 5,  # 地热井数目/个
        'p_fc_max': 300,  # 燃料电池容量/kw
        'p_hpg_max': 300,  # 地源热泵功率/kw
        'p_hp_max': 300,  # 空气源热泵功率/kw
        'p_eb_max': 300  # 电热锅炉功率/kw
        'p_el_max': 300,  # 电解槽功率/kw
        'hst': 100,  # 储氢罐容量/kg
        'm_ht': 100,  # 储热罐/kg
        'm_ct': 100,  # 冷水罐/kg
        'area_pv': 600,  # 光伏面积/m2
        'area_sc': 500,  # 集热器面积/m2
        'p_co': 300,  #氢压机功率/kw

        "equipment_cost": 1000,  #设备总投资/万元
        "receive_year": 10,  # 投资回报年限/年
}
#离网结合运行后的输出
itgrid_operation_output_json = {
        "operation_cost": 10,  # 年化运行成本/万元
        "cost_save_rate": 20,  #运行成本节约比例%
        "co2":5,  #总碳排/t
        "cer":20,  #碳减排量
        "cer_rate": 70,  # 碳减排率%
        "cer_perm2":2  #每平米的碳减排量/t
}


```