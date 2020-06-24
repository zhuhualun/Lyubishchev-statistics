# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 22:28:40 2020

@author: 56932
"""

import os
import time
import datetime
import pandas as pd
import numpy as np
from collections import defaultdict
import matplotlib as plt
import matplotlib.pyplot as plt
print("请把excel文件放在%s下"%os.getcwd())

data = pd.read_excel('每日工作计时.xlsx')
data['日期'] = data['日期'].fillna(method='ffill')
data = data.set_index('日期')

plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 魔法函数
# %matplotlib inline

"""
00 获取某一格的时间,
"""
def get_time(data,col,i):
    time_index = data_i.index.values
    date = str(time_index[0])[:10]
    t_y = int(date[:4])
    t_m = int(date[5:7])
    t_d = int(date[8:10])
    time = str(data[col][i])
    t_h = int(time[:2])
    t_min = int(time[3:5])
    t_s = int(time[6:8])
    t = (t_y,t_m,t_d,t_h,t_min,t_s,0,0,0)
    return t

"""
01 获取一周总表
"""
def get_week(data):
    today = datetime.date.today()
    weekday = datetime.datetime.now().weekday()
    t1 = today - datetime.timedelta(7+weekday)
    t2 = today - datetime.timedelta(1+weekday)
    data_i = data[t1:t2]
    data_i = data_i.dropna(axis=0,how='any',subset=['开始时间','结束时间'])
    return (data_i,t1,t2)

data_i = get_week(data)[0]
print(data_i)

"""
02 定位内容属于哪一列
"""
# 返回列的编号以及单元格的内容
def get_col_posi(data_i,loc):
    for i in range(2,data_i.shape[1]):
        data_col_i = data_i.iloc[loc][i]
        if type(data_col_i) is str:
            if len(data_col_i)>0:
                j = i
                #print(data_col_i)
                return (j,data_col_i)
            
"""
03 计算时间差，并储存
"""
# 返回各项的时间
def get_cost(data):
    cost_lst = []
    item_lst = []
    for i in range(data_i.shape[0]):
        dic = {}
        dic_item = {}
        p = get_col_posi(data_i,i)[0]
        t1 = get_time(data_i,'开始时间',i)
        t2 = get_time(data_i,'结束时间',i)
        t1_f = time.mktime(t1)
        t2_f = time.mktime(t2)
        col_name_lst = data_i.columns.values.tolist()
        item = get_col_posi(data_i,p)[1]
        col_name_i = col_name_lst[p]
        #print(col_name_i)
        if t1_f < t2_f:
            delte_t = (t2_f - t1_f)/3600
        else:
            delte_t = (t2_f - t1_f + 3600*24)/3600
        dic[col_name_i] = round(delte_t,1)
        dic_item[item] = round(delte_t,1)
        cost_lst.append(dic)
        item_lst.append(dic_item)
    return(cost_lst,item_lst)


"""
时间累加
"""
def sum_dict(a,b):
    temp = dict()
    for key in a.keys()| b.keys():
        temp[key] = sum([d.get(key, 0) for d in (a, b)])
    return temp


"""
入口
"""
if __name__ == '__main__':
    cost_lst = get_cost(data_i)[0]
    item_lst = get_cost(data_i)[1]
    # 计算大项
    dic = {}    
    for i in cost_lst:
        dic = sum_dict(dic,i)
    
    # 计算小项
    dic_item = {}
    for i in item_lst:
        dic_item = sum_dict(dic_item,i)
    
    cost_df = pd.DataFrame(pd.Series(dic),columns=['用时(h)'])
    item_df = pd.DataFrame(pd.Series(dic_item),columns=['用时(h)'])

    # 保存图片
    # 获取日期
    date1 = get_week(data)[1]
    date2 = get_week(data)[2]
    b_tm_year = date1.timetuple()[0]
    b_tm_mon = date1.timetuple()[1]
    b_tm_day = date1.timetuple()[2]
    e_tm_year = date2.timetuple()[0]
    e_tm_mon = date2.timetuple()[1]
    e_tm_day = date2.timetuple()[2]
    
    # 文件名和路径
    filedate1 = F"\\周报大项{b_tm_year}年{b_tm_mon}月{b_tm_day}日_{e_tm_mon}月{e_tm_day}日"
    filedate2 = F"\\周报小项{b_tm_year}年{b_tm_mon}月{b_tm_day}日_{e_tm_mon}月{e_tm_day}日"
    foldername = "柳比歇夫工作法"
    if not os.path.exists(foldername):
        os.makedirs(foldername)
    filepath = os.getcwd() + "\\" + foldername
    filename1 = filepath + filedate1
    filename2 = filepath + filedate2

    # 900*500像素
    plt.rcParams['figure.figsize'] = (9.0,5.0)
    plt.rcParams['savefig.dpi'] = 100
    cost_df.plot(kind='barh')
    plt.savefig("%s.png"%filename1)
    plt.show()
    
    plt.rcParams['figure.figsize'] = (9.0,5.0)
    plt.rcParams['savefig.dpi'] = 100
    item_df.plot(kind='barh')
    plt.savefig("%s.png"%filename2)
    plt.show()

    print("文件已保存在%s"%filepath)
    
"""
个人微信公众号：《纸箱之神》，不定期分享关于自律、自我管理、互联网运营等知识，欢迎关注。
"""
    

    
    
