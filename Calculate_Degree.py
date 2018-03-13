import datetime
from collections import OrderedDict
import xlrd
import sys

data = xlrd.open_workbook('geocentric 1990-2035.xls')
PLANET_NUM = 10


def generator_dict(data, planet):

    table = data.sheet_by_index(0)
    rows = table.nrows
    ordered_dict = OrderedDict()
    count = 0
    pre_degree = table.cell(2,planet[0]).value #数据从第3行开始

    for x, y in zip(table.col_values(0), table.col_values(planet[0])): #col_values(0)日期
        #判断是否逆行
        if table.cell(count, 0).ctype == 3:  # 判断是否为日期格式
            if y - pre_degree < 0 and abs(pre_degree - y) < 180: #判断是否逆行，(pre_degree - y)排除从0度开始的情况，第一次pre_diff为0不准确，待改进
                tup = y, "R"
            elif y - pre_degree >180:   #判断逆行时从0度回退到360度的情况
                tup = y, "R"
            else:
                tup = y, "A"
            pre_degree = y

        if count < rows:
            if table.cell(count, 0).ctype == 3: #判断是否为日期格式
                date_value= xlrd.xldate.xldate_as_datetime(table.cell(count,0).value, 0)  #将EXCEL中的日期直接转化为datetime对象
                format_date = date_value.strftime('%Y-%m-%d')   #带时间('%y-%m-%d %I:%M:%S')
                ordered_dict[format_date] = tup
            else:
                count = count + 1
                continue
        count = count + 1

    return ordered_dict



def find_input_degree(planet_dict,deviation):
    #输入要查到的度数
    print("input %s degree you want to find" % input_value)
    input_degree = input(":")
    degree = float(input_degree)
    try:
        float(degree);
    except(ValueError):
        print("%s it's not a number" % num_float);

    result_dict = OrderedDict()
    temp = OrderedDict()
    planet_key = list(planet_dict.keys())
    count=0

    #判断第一个值是否为逆行
    first_diff = planet_dict[planet_key[0]][0] - planet_dict[planet_key[1]][0]
    if first_diff < 0 :  #如果第一天小于第二天，并且不是从新开始
        pre_v = planet_dict[planet_key[0]][0] - 1   #从EXLCE第一行开始计算
        planet_stat = 'A'
    else:
        pre_v = planet_dict[planet_key[0]][0] + 1  # 从EXLCE第一行开始计算
        planet_stat = 'R'

    #把各种不同周期数据放入有序字典，取最接近输入值的数字
    for k, v in planet_dict.items():
        count = count + 1
        print(count)
        diff = abs(pre_v - v[0])       #当前日期与上个日期度数差值
        if diff < 180 and v[1] == planet_stat:      #确保当度数回到0时能确清空temp字典，重新计算下一个0-360度
            temp[k] = abs(v[0] - degree)        #放到0-360度集合
            pre_v = v[0]
        else:
            if temp:    #temp是否为空
                min_tuple = min(temp.items(),key = lambda x:x[1])     #diff>180 一圈地完成，计算与input_degree最近值
                if min_tuple[1] < float(deviation[2]):  #判断要查找的值是否小于行星最大日移动度数
                    result_dict[min_tuple[0]] = planet_dict[min_tuple[0]]       #写入结果
            temp.clear()         #清空本周期数据
            pre_v = 180
            planet_stat = v[1]

    return result_dict


if __name__ == "__main__":

    deviation = {'sun': (1, 'sun', 1),
                 'mercury': (2,'mercury', 4.09),
                 'venus': (3,'venus', 1.6),
                 'mars': (4,'mars', 0.53),
                 'jupiter': (5,'jupiter', 0.083),
                 'saturn': (6,'saturn', 0.335),
                 'uranus': (7,'uranus', 0.0117),
                 'neptune': (8,'neptune', 0.006),
                 'pluto': (9,'pluto', 0.004),
                 'moon': (10,'moon', 13.177)}
    while True:
        val = None
        input_value = input("input \"quit\" to exit \n"
                            "input planet name:"
                            ).lower()
        if input_value in deviation.keys():
            init_dict = generator_dict(data, deviation[input_value])
            result_dict = find_input_degree(init_dict,deviation[input_value])
            print("result %s degree is:" % input_value)
            print(result_dict)
        elif input_value == "quit":
            sys.exit()
        else:
            print("Re-enter the planet name!")
        continue
        print(numbers_to_planet(val))
