#
import pandas as pd
import numpy as np
#读取数据
#FBA后台退款表
fba_refund = pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\5月-FBA退款订单.xlsx",dtype=str)
#FBM后台退款表
fbm_refund = pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\5月-FBM退款订单.xlsx",dtype=str)
#asinkingFBA退款表中的退款原因
asinking_fba_refund = pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\德国-asinking-退货订单.xlsx",dtype=str,)

#亚马逊后台退款原因匹配表
reasons_refund_parameter_1 =pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\FBA亚马逊退款原因.xlsx",dtype=str,sheet_name='工作表1')
reasons_refund_parameter_2 =pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\FBA亚马逊退款原因.xlsx",dtype=str,sheet_name='工作表2')
#人工维护订单退款原因
reasons_refund_artificial =pd.read_excel(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\德国-5月售后.xlsx",dtype=str)
#============================================================================================================================================================
#合并FBA退款原因
fba_refund_merge = pd.merge(fba_refund,reasons_refund_artificial,how='left',left_on=('订单编号','易仓sku'),right_on=('订单号','SKU'))
#合并FBM退款原因
fbm_refund_merge = pd.merge(fbm_refund,reasons_refund_artificial,how='left',left_on=('订单编号','易仓sku'),right_on=('订单号','SKU'))

asinking_fba_refund_merge = pd.merge(asinking_fba_refund,reasons_refund_parameter_1,how='left',left_on=('退货原因'),right_on=('退款原因'))
asinking_fba_refund_merge_2 = pd.merge(asinking_fba_refund_merge,reasons_refund_parameter_2,how='left',left_on=('退款原因.1'),right_on=('退款原因'))




#删除无需字段
Refund_form = asinking_fba_refund_merge_2.drop(asinking_fba_refund_merge_2.iloc[:, [-1,-2,-3,-4,-5,-6,-7,-8,-9,-10,-11,-12,-13,-14,-15,-16,-17,-18,-19,-20,-21,-22,-23,-26,-27,-28,-29,-30,-31,-32,-33,-34,-35,-36,-37,-38,-39,-40,-41,-42,-43]],axis=1)
fba_refund_merge_1 = pd.merge(fba_refund_merge,Refund_form,how='left',left_on=('订单编号','易仓sku'),right_on=('订单号','sku'))
#
fba_refund_merge_1.fillna('0',inplace=True)
fba_refund_merge_2 = fba_refund_merge_1.drop(fba_refund_merge_1.iloc[:, [-3,-4,-5,-6,-7,-8,-9,-11,-12,-13,-14,-15,-16,-17,-18,-19,-20,-21,-22,-23,-24,-25,-26,-27,-28,-29,-30,-31,-32,-33,-34,-35,-36,-37,-38,]],axis=1)


fba_refund_merge_2['退款原因_y'] = np.where(fba_refund_merge_2['退款原因_y']=='0',fba_refund_merge_2['退款原因.1_y'],fba_refund_merge_2['退款原因_y'])
fba_refund_merge_2['退款类型_y'] = np.where(fba_refund_merge_2['退款类型_y']=='0',fba_refund_merge_2['退款类型'],fba_refund_merge_2['退款类型_y'])
fba_refund_merge_2['分类依据提示_y'] = np.where(fba_refund_merge_2['分类依据提示_y']=='0',fba_refund_merge_2['分类依据提示'],fba_refund_merge_2['分类依据提示_y'])

#
fba_refund_merge_3 = fba_refund_merge_2.drop(fba_refund_merge_2.iloc[:, [-1,-2,-3,-4,-9,-10,-11,-12,-13,-14,-15,-16,-17]],axis=1)
fbm_refund_merge_4 = fbm_refund_merge.drop(fbm_refund_merge.iloc[:, [-1,-2,-3,-4,-5,-6,-7,-8,-9,-14,-15,-16,-17,-18,-19,-20,-21,-22]],axis=1)
#=============================================================================================================================================================
# #输出结果
with pd.ExcelWriter(r"C:\Users\huawei\Desktop\代码\pythonProject1\项目\马倩佳退款\5月德国退款汇总表（马倩佳）.xlsx") as writer:
     fba_refund_merge_3.to_excel(writer,sheet_name='FBA退款', index=False)
     fbm_refund_merge_4.to_excel(writer,sheet_name='FBM退款', index=False)


# #==========================================================



