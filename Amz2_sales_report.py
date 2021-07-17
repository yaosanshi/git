import pandas as pd
import numpy as np

#读取销售报表
Sales_details = pd.read_excel(r"E:\python\项目\李青销售报表\订单列表20210701-36-269158797032677376-0.xlsx")
Comparison_table = pd.read_excel(r"E:\python\项目\李青销售报表\对照表.xlsx")#读取对照表
lisiting = pd.read_excel(r"E:\python\项目\李青销售报表\Listing20210701-269165970999595008-0.xlsx")#读取listing列表
#
Sales_details['订单币种'].replace('GBP', 9,inplace = True)#币种变更为汇率
Sales_details['订单币种'].replace('CAD', 5,inplace = True)#币种变更为汇率
Sales_details['订单币种'].replace('EUR', 7.7,inplace = True)#币种变更为汇率
Sales_details['订单币种'].replace('JPY', 0.058,inplace = True)#币种变更为汇率
Sales_details['订单币种'].replace('USD', 6.3,inplace = True)#币种变更为汇率
Sales_details['订购日期'] = Sales_details['订购日期'].astype("str").str[0:10] #截取年-月，变更日期格式，方便透视
#报表合并
merge_Sales_details = pd.merge(
                                Sales_details,Comparison_table,
                                how='left',
                                left_on=('ASIN','MSKU'),
                                right_on=('ASIN','平台SKU'))#合并对照表
lisiting_order = pd.merge(merge_Sales_details,lisiting,
                          how='left',
                          left_on=('ASIN','MSKU','店铺'),
                          right_on=('ASIN','MSKU','店铺'))#合并Listing报表，取销售额
lisiting_order['总销售额'] = lisiting_order['订单币种']*lisiting_order['价格']*lisiting_order['数量']#计算总销售额
perspective = pd.pivot_table(lisiting_order,
                             index=[u'订购日期',u'销售人员',u'ASIN',u'MSKU',u'品类'],
                             values=(u'总销售额',u'数量'),
                             aggfunc=[np.sum])#数据透视取值结果
perspective.to_excel(r"E:\python\项目\李青销售报表\结果\亚马逊2部销售分析表.xlsx")#导出结果