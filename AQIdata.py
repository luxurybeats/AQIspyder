# -*- coding:utf-8 -*-
"""

使用python爬取http://www.cdepb.gov.cn/cdepbws/web/gov/airquality.aspx 数据并保存到Excel

"""

__author__ = 'Luxury'

import requests
import xlwt, xlrd
import re
import numpy as np
import lxml.etree
import xlutils.copy

# 初始化所用列表
DATA = []                   # 实时数据
DayData = []                # 每日数据
Index = ['AQI','SO2','NO2','PM10','CO2','O3','PM2.5']           # 空气质量指标
DayDataXcagtegories=["青羊区","金牛区","锦江区","武侯区","成华区","高新区", "龙泉驿区",       # 成都地区
                     "青白江区","新都区","温江区","双流区","郫县","天府新区","都江堰市",
                     "崇州市","新津县","彭州市","邛崃市","大邑县","蒲江县"]
# 获取整个网页
first_url = 'http://www.cdepb.gov.cn/cdepbws/web/gov/airquality.aspx'
r = requests.get(first_url)
html = r.text

# 获取当前时间
data_time = lxml.etree.HTML(html).xpath('//*[@id="ContentBody_AQITime"]/text()')
data_time = str(data_time).replace(' ','').replace('[u\'','').replace('\']','').replace\
    ('\u5e74','.').replace('\u6708','.').replace('\u65e5','.').replace('\u65f6','')

# 获取地区记录时间
data_time2 = re.findall(r'inline-block;" value.*? />', html, re.S)[0]
data_time2 = re.findall(r'value.*/>',data_time2)[0]
data_time2 = re.findall(r'".*"',data_time2)[0].replace('\"','')

# 使用xlrd 与 xlutils.copy 对数据进行追加操作
dayDatabookopen = xlrd.open_workbook(r'AQI.xls')
sheet = dayDatabookopen.sheet_by_index(0)
col = sheet.col_values(0)                       # 获取第一列
newdayDatabookopen = xlutils.copy.copy(dayDatabookopen)
newsheet = newdayDatabookopen.get_sheet(0)

# 使用xlwt 设定编码方式
workbook=xlwt.Workbook(encoding='utf-8')
dayDatabook=xlwt.Workbook(encoding='utf-8')

# 制地区AQI数据表
daybooksheet=dayDatabook.add_sheet('sheet', cell_overwrite_ok=False)

# 制实时数据地区表
booksheet1=workbook.add_sheet('成都市', cell_overwrite_ok=True)
booksheet2=workbook.add_sheet('君平街', cell_overwrite_ok=True)
booksheet3=workbook.add_sheet('大石西路', cell_overwrite_ok=True)
booksheet4=workbook.add_sheet('梁家巷', cell_overwrite_ok=True)
booksheet5=workbook.add_sheet('金泉两河', cell_overwrite_ok=True)
booksheet6=workbook.add_sheet('沙河滩', cell_overwrite_ok=True)
booksheet7=workbook.add_sheet('三瓦窑', cell_overwrite_ok=True)
booksheet8=workbook.add_sheet('十里店', cell_overwrite_ok=True)
booksheet9=workbook.add_sheet('灵岩山', cell_overwrite_ok=True)

# 获取地区实时数据
content = re.findall(r'monitorChartJson=.*diviChartJson=', html, re.S )[0]
content = re.findall(r'{.*}', content)[0].replace('null','\'null\'')
content = re.findall(r'rows:.*?]}]}', content)
for i in content:
    content = re.findall(r'{.*}', i)[0]
    content = re.findall(r'data:.*?}]}', content)
    for j in content:
        content =  re.findall(r'{.*}', j)[0].replace(']}','')
        content = eval(content)         # 字符串转化为字典类型
        for k in content:
            DATA.append(k['y'])

# 整理数据为多维数组
DATA = np.array(DATA).reshape(9,7,24)

# 获取地区每日数据
DayDataChart = re.findall(r'DayDataChartJson=.*DayDataXcagtegories=', html, re.S)[0]
DayDataChart = re.findall(r'{.*}', DayDataChart)[0].replace('null','\'null\'')
DayDataChart = re.findall(r'data:.*],dataLabels', DayDataChart)[0]
DayDataChart = re.findall(r'{.*}', DayDataChart)[0].replace('y','\'y\'').replace('color','\'color\'')
DayDataChart = eval(DayDataChart)
for i in DayDataChart:
    DayData.append(i['y'])

# 写入表格每日程序
for i in range(len(DayDataXcagtegories)):
    newsheet.write(0,i+1,DayDataXcagtegories[i])
    newsheet.write(len(col)+1,0,data_time2)
    newsheet.write(len(col)+1,i+1,DayData[i])

# 写入表格实时数据 为四维数据
for i in range(7):
    for j in range(24):
        booksheet1.write(0,j+1,j)
        booksheet1.write(i+1,0,Index[i])
        booksheet1.write(i+1,j+1,DATA[0][i][j])

for i in range(7):
    for j in range(24):
        booksheet2.write(0,j+1,j)
        booksheet2.write(i+1,0,Index[i])
        booksheet2.write(i+1,j+1,DATA[1][i][j])

for i in range(7):
    for j in range(24):
        booksheet3.write(0,j+1,j)
        booksheet3.write(i+1,0,Index[i])
        booksheet3.write(i+1,j+1,DATA[2][i][j])

for i in range(7):
    for j in range(24):
        booksheet4.write(0,j+1,j)
        booksheet4.write(i+1,0,Index[i])
        booksheet4.write(i+1,j+1,DATA[3][i][j])

for i in range(7):
    for j in range(24):
        booksheet5.write(0,j+1,j)
        booksheet5.write(i+1,0,Index[i])
        booksheet5.write(i+1,j+1,DATA[4][i][j])

for i in range(7):
    for j in range(24):
        booksheet6.write(0,j+1,j)
        booksheet6.write(i+1,0,Index[i])
        booksheet6.write(i+1,j+1,DATA[5][i][j])

for i in range(7):
    for j in range(24):
        booksheet7.write(0,j+1,j)
        booksheet7.write(i+1,0,Index[i])
        booksheet7.write(i+1,j+1,DATA[6][i][j])

for i in range(7):
    for j in range(24):
        booksheet8.write(0,j+1,j)
        booksheet8.write(i+1,0,Index[i])
        booksheet8.write(i+1,j+1,DATA[7][i][j])

for i in range(7):
    for j in range(24):
        booksheet9.write(0,j+1,j)
        booksheet9.write(i+1,0,Index[i])
        booksheet9.write(i+1,j+1,DATA[8][i][j])

# 以时间为表格名保存实时数据
workbook.save('%s.xls' % data_time)

# 保存每日数据
newdayDatabookopen.save('AQI.xls')

# 显示是否上传成功
print(u'%s 数据已上传至表格' % data_time)







