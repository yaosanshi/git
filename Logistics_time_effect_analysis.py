import datetime
import pandas as pd
import numpy as np

#读取文件后分组聚合
prescription = pd.read_excel(r'E:\python\项目\物流部需求\2021-6 海外仓时效分析表7.1-7.8.xlsx')#读取文件
daily_sales_group_data_1 = pd.pivot_table(prescription,
                                            index=[u'仓库代码',u'出货时间',u'当前状态'],
                                            values=[u'耗时天数',u'跟踪号'], aggfunc={'耗时天数': sum,'跟踪号': 'count'},
                                            margins=True, margins_name='总计')
daily_sales_group_data_1.to_csv(r'E:\python\项目\物流部需求\测试.csv')
#表格重塑
table = pd.read_csv(r'E:\python\项目\物流部需求\测试.csv')
remodeling = table.pivot_table(index=['仓库代码','出货时间'], columns='当前状态',aggfunc=np.sum,margins=True)#重塑
remodeling.to_csv(r'E:\python\项目\物流部需求\分析报表.csv')


xinbiao = pd.read_csv(r'E:\python\项目\物流部需求\分析报表.csv')
xinbiao1 = xinbiao.drop(xinbiao.iloc[:, [2,3,4]],axis=1)#删除多余列
xinbiao1.columns = ['仓库代码','发货时间','总时效','可能异常','成功签收','运输途中','发货数量']#暴力更改列头
xinbiao1.drop([1],axis=0,inplace=True)#删除第二列
xinbiao1.drop([0],axis=0,inplace=True)#删除第一列
xinbiao1.to_excel(r'E:\python\项目\物流部需求\物流时效分析报表.xlsx')
xinbiao2 = pd.read_excel(r'E:\python\项目\物流部需求\物流时效分析报表.xlsx')
xinbiao2['平均时效'] = round(xinbiao2['总时效']/xinbiao2['发货数量'],2)#计算时效后保留2位小数
xinbiao2['签收率'] = xinbiao2['成功签收']/xinbiao2['发货数量']#计算签收率
# xinbiao2['运输途中（总）'] = xinbiao2['到达待取']+xinbiao2['运输途中']#计算签收率
# xinbiao2['派送异常'] = xinbiao2['可能异常']+xinbiao2['投递失败']#计算签收率
xinbiao2.drop(xinbiao2.iloc[:, [0,3]],axis=1,inplace=True)#删除多余列

#通过数据透视合并仓库代码单元格
p = pd.pivot_table(xinbiao2,
                            index=[u'仓库代码',u'发货时间'],
                            values=[u'运输途中',u'可能异常',u'成功签收',u'发货数量',u'平均时效',u'签收率',],
                            margins=True)
#转换百分比
p['签收率'] = p['签收率'].apply(lambda x: format(x, '.2%'))
#交换列位置
cols = p.columns.tolist()
cols.insert(cols.index('可能异常'), cols.pop(cols.index('运输途中')))
cols.insert(-1, cols.pop(cols.index('平均时效')))
df_final = p[cols]
#生成最终报表
df_final.to_excel(r'E:\python\项目\物流部需求\物流时效分析报表（最新）.xlsx')




