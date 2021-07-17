import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
# def warehousing_form():
#上架报表
shelves = pd.read_excel(r"E:\python\项目\监控报表项目\上架数据\PUT-S210716184134.xls",dtype=str)#读取上架数据
product = pd.read_excel(r"E:\python\项目\监控报表项目\上架数据\产品明细.xlsx")#读取产品资料并解码中文encoding='GB18030'
num_shelvse = shelves['仓库'].str.split(',').to_list()#输出仓库列为列表
for i in num_shelvse:#遍历仓库区分海外与国内
    if i ==['中转WINIT/4PX/YKD']:
        shelves.replace([i],'国内仓',inplace = True)
    elif i ==['深圳仓']:
        shelves.replace([i], '国内仓', inplace=True)
    elif i ==['CD中转']:
        shelves.replace([i], '国内仓', inplace=True)
    elif i ==['中转UK8']:
        shelves.replace([i], '国内仓', inplace=True)
    elif i ==['中转至FBA']:
        shelves.replace([i], '国内仓', inplace=True)
    elif i ==['宁波中转']:
        shelves.replace([i], '国内仓', inplace=True)
    else:
        shelves.replace([i], '海外仓', inplace=True)
price = pd.merge(shelves,product,
                 how='left',
                 left_on=('SKU'),
                 right_on=('产品SKU'))#两表合并
price.to_csv(r"E:\python\项目\监控报表项目\上架数据\上架表.csv")#输出为CSV--excel格式复杂某些字段无法计算
#把时间字符串转为时间类型设置为索引
num =pd.read_csv(r"E:\python\项目\监控报表项目\上架数据\上架表.csv")
num['上架成本']=num['最近单价']*num['上架数量']#新增成本列

num['上架时间'] = num['上架时间'].astype("str").str[0:10]
pd.set_option('display.float_format',lambda x : '%.3f' % x)#不显示科学计数法

num_1 = pd.pivot_table(num,
                       index=['上架时间','仓库'],
                       values=['上架成本'],
                       aggfunc={'上架成本':np.sum})
print(num_1)#打印上架成本

#==========================================================================================================================================

#库存报表
#计算国内库存在途&在库成本
overseas = pd.read_csv(r"E:\python\项目\监控报表项目\库存数据\bdsmPRODUCT_SALES_ANALYSIS_REPORT_30922.csv",low_memory=False,dtype=str)#读原表
overseas.drop(overseas.index[:2], inplace=True)
overseas.to_csv(r"E:\python\项目\监控报表项目\库存数据\国内仓.csv",header=None)#原表格式异常，输出CSV重新处理
overseas_1 = pd.read_csv(r"E:\python\项目\监控报表项目\库存数据\国内仓.csv",low_memory=False)#读新CSV

# overseas1 = overseas.drop(index=[0],axis=0)
overseas_1['在途成本'] = overseas_1['在途']*overseas_1['默认采购价(RMB)']
purchasing_cost = overseas_1['在途成本'].sum()
inventory_cost = overseas_1['库存成本(RMB)'].sum()
print('国内仓在途成本为：{}'.format(purchasing_cost))
print('国内仓在库成本为：{}'.format(inventory_cost))


#计算海外仓在途&在库成本
overseas2 = pd.read_csv(r"E:\python\项目\监控报表项目\库存数据\bdsmPRODUCT_SALES_ANALYSIS_REPORT_30923.csv",low_memory=False,dtype=str)#读原表
overseas2.drop(overseas2.index[:2], inplace=True)
overseas2.to_csv(r"E:\python\项目\监控报表项目\库存数据\海外仓.csv",header=None)#原表格式异常，输出CSV重新处理
overseas_2 = pd.read_csv(r"E:\python\项目\监控报表项目\库存数据\海外仓.csv",low_memory=False)#读新CSV

# # overseas1 = overseas.drop(index=[0],axis=0)
overseas_2['在途成本'] = overseas_2['在途']*overseas_2['默认采购价(RMB)']
purchasing_cost_2 = overseas_2['在途成本'].sum()
inventory_cost_2 = overseas_2['库存成本(RMB)'].sum()
print('海外仓在途成本为：{}'.format(purchasing_cost_2))
print('海外仓在库成本为：{}'.format(inventory_cost_2))

#==============================================================================================================================================================

#FBA销售额报表
lisiting = pd.read_excel(r"E:\python\项目\监控报表项目\销售数据\Listing20210715-274244389069930497-0.xlsx")#读取Lisiting
order = pd.read_excel(r"E:\python\项目\监控报表项目\销售数据\订单列表20210716-97-274610194383224832-0.xlsx")#读取asinking销售报表
#销售报表与listing合并（获取销售额）
lisiting_order = pd.merge(order,lisiting,
                          how='left',
                          left_on=('ASIN','MSKU','店铺'),
                          right_on=('ASIN','MSKU','店铺'))
#合并产品资料（获取成本）
lisiting_order_cost =pd.merge(lisiting_order,product,
                              how='left',
                              left_on=('SKU_x'),
                              right_on=('产品SKU'))
#判断币种，插入汇率
num_lisiting_order_cost = lisiting_order_cost['订单币种'].str.split(',').to_list()

for i in num_lisiting_order_cost:
    if i == ['EUR']:
        lisiting_order_cost.replace([i], 7.7, inplace=True)
    elif i == ['CAD']:
        lisiting_order_cost.replace([i], 5.2, inplace=True)
    elif i == ['GBP']:
        lisiting_order_cost.replace([i], 9, inplace=True)
    elif i == ['USD']:
        lisiting_order_cost.replace([i], 6.3, inplace=True)
    elif i == ['JPY']:
        lisiting_order_cost.replace([i], 0.058, inplace=True)
lisiting_order_cost['总销售额'] = lisiting_order_cost['订单币种']*lisiting_order_cost['价格']*lisiting_order_cost['数量']

# lisiting_order_cost['总成本']= lisiting_order_cost['数量']*lisiting_order_cost['最近单价']
#此处存在问题，直接计算会导致结果错误，暂未解决
lisiting_order_cost.to_csv(r"E:\python\项目\监控报表项目\销售数据\FBA销售额.csv")

lisiting_order_cost_1 = pd.read_csv(r"E:\python\项目\监控报表项目\销售数据\FBA销售额.csv")
lisiting_order_cost_1['总成本']= lisiting_order_cost_1['数量']*lisiting_order_cost_1['最近单价']
# a.to_csv(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\监控报表项目\销售数据\销售额_1.csv")

lisiting_order_cost_1['订购日期'] = lisiting_order_cost_1['订购日期'].astype("str").str[0:10]

num_2 = pd.pivot_table(lisiting_order_cost_1,
                       index=['订购日期'],
                       values=['总销售额','最近单价'],
                       aggfunc=[np.sum])
print(num_2)#

#=======================================================================================================
#自发货销售报表
sales_list =pd.read_csv(r"E:\python\项目\监控报表项目\销售数据\bdsmsales_wh_30926.csv",dtype=str,low_memory=False)
sales_list.drop(sales_list.index[:5], inplace=True)
sales_list.to_csv(r"E:\python\项目\监控报表项目\销售数据\易仓销售额.csv",header=None)#输出后删除表头

sales_list_1 = pd.read_csv(r"E:\python\项目\监控报表项目\销售数据\易仓销售额.csv",low_memory=False)
sales_list_1 = sales_list_1[ ~ sales_list_1['订单类型'].str.contains('refund','resend')]#删除重发退款订单
sales_list_1 = sales_list_1[ ~ sales_list_1['订单类型'].str.contains('resend')]
sales_list_1 = sales_list_1[ ~ sales_list_1['发运仓库'].str.contains('FBA')]#删除FBA订单
sales_list_1['总采购成本'] = sales_list_1['采购运费']+sales_list_1['采购税费']+sales_list_1['采购成本']
sales_list_1['付款时间'] = sales_list_1['付款时间'].astype("str").str[0:10]

sales_list_1.to_csv(r"E:\python\项目\监控报表项目\销售数据\易仓销售额_1.csv",)


num_3 = pd.pivot_table(sales_list_1,
                       index=['付款时间'],
                       values=['订单总金额(包含客户运费、平台补贴)','总采购成本'],
                       aggfunc=[np.sum])
print(num_3)#打印上架成本
# #==============================================================================================================================================
#需优化内容
#1自动下载各报表
#2自动输入内容到库存监控报表
#3下载报表分时间统计