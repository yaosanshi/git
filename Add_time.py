import pandas as pd
import os
import datetime
#1、通过表名获取下载日期，填充到表格中去
#2、取下载日期减去1天，作为数据日期，填充到表格中
filename = os.listdir(r"E:\python\项目\批量修改报表\export_files_clean")
for i in filename:
    Download_time = i[15:21]#根据表明获取表格下载时间
    Download_time1 = pd.to_datetime(Download_time, format='%y%m%d')
    Download_time2 = Download_time1 + datetime.timedelta(-1)#数据时间
    df = pd.read_csv(r"E:\python\项目\批量修改报表\export_files_clean\\" + i, encoding='gbk', low_memory=False)
    df['下载时间']= Download_time1
    df['数据时间']= Download_time2
    df.to_csv(r"E:\python\项目\批量修改报表\save\\" + i)
print('修改完毕')