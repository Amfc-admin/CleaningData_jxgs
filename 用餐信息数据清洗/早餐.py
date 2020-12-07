import datetime
import os

import pandas as pd

# 修改工作路径,os.getcwd()为让前路径,os.chdir为修改之后的路径
os.getcwd()
path = 'F:\#!Python_porject\用餐信息数据清洗\数据导出'
os.chdir(path)
dir_a = os.getcwd()
print('当前路径为: ' + dir_a)

filelist = []
for root, dirs, files in os.walk(path):
    for file in files:
        if os.path.splitext(file)[1] == '.xlsx':
            filelist.append(file)

# print(filelist)

pd_list = []
for i in range(len(filelist)):
    pd_list.append(pd.read_excel(filelist[i]), )

merge = pd.concat(pd_list)
# print(All)
merge.to_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx', sheet_name='All', index=False)
print('数据合并完成')

sort1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')
sort1.sort_values(by=['所属部门', '用餐类型'], inplace=True, ascending=[True, True])
# print(px)
# sort1.to_excel('F:\#!Python_porject\用餐信息数据清洗\总表\A_总表.xlsx', sheet_name='排序部门')
# print('用餐部门排序完成')

sort2 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')
sort2.sort_values(by=['用餐类型', '所属部门'], inplace=True, ascending=[True, True])
# print(px)
# sort2.to_excel('F:\#!Python_porject\用餐信息数据清洗\总表\A_总表.xlsx', sheet_name='排序用餐')
# print('用餐类型排序完成')


screen_read = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')

screen1 = screen_read[(screen_read[u'用餐类型'] == '早餐')]
count_screen1 = screen1.shape[0]
# print(count_screen1)

date_now = str(datetime.date.today())
date_tomorrow = (datetime.date.today() + datetime.timedelta(days=+1)).strftime('%Y-%m-%d')
put1 = pd.Series([date_tomorrow], index=[1], name='日期')
put2 = pd.Series(['早餐'], index=[1], name='用餐类型')
put3 = pd.Series([count_screen1], index=[1], name='合计人数')
count_1 = pd.DataFrame({put1.name: put1, put2.name: put2, put3.name: put3})  # 用DataFrame要写成字典形式

####################################################################################################


screen_read2 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')
screen4 = screen_read2[(screen_read2[u'所属部门'] == '生产科')]
count_screen4 = screen4.shape[0]

screen5 = screen_read2[(screen_read2[u'所属部门'] == '经营科')]
count_screen5 = screen5.shape[0]

screen6 = screen_read2[(screen_read2[u'所属部门'] == '党委办公室')]
count_screen6 = screen6.shape[0]

screen7 = screen_read2[(screen_read2[u'所属部门'] == '安全环保科')]
count_screen7 = screen7.shape[0]

screen8 = screen_read2[(screen_read2[u'所属部门'] == '技术监督科')]
count_screen8 = screen8.shape[0]

screen9 = screen_read2[(screen_read2[u'所属部门'] == '企管科')]
count_screen9 = screen9.shape[0]

screen10 = screen_read2[(screen_read2[u'所属部门'] == '容器制造分公司')]
count_screen10 = screen10.shape[0]

screen11 = screen_read2[(screen_read2[u'所属部门'] == '修理分公司')]
count_screen11 = screen11.shape[0]

screen12 = screen_read2[(screen_read2[u'所属部门'] == '光正分公司')]
count_screen12 = screen12.shape[0]

screen13 = screen_read2[(screen_read2[u'所属部门'] == '自控维护分公司')]
count_screen13 = screen13.shape[0]

screen14 = screen_read2[(screen_read2[u'所属部门'] == '设备制造分公司')]
count_screen14 = screen14.shape[0]

screen15 = screen_read2[(screen_read2[u'所属部门'] == '钻采设备分公司')]
count_screen15 = screen15.shape[0]

screen16 = screen_read2[(screen_read2[u'所属部门'] == '油管检修分公司')]
count_screen16 = screen16.shape[0]

screen17 = screen_read2[(screen_read2[u'所属部门'] == '防腐分公司')]
count_screen17 = screen17.shape[0]

screen18 = screen_read2[(screen_read2[u'所属部门'] == '燃烧器项目组')]
count_screen18 = screen18.shape[0]

screen19 = screen_read2[(screen_read2[u'所属部门'] == '销售部')]
count_screen19 = screen19.shape[0]

screen20 = screen_read2[(screen_read2[u'所属部门'] == '生产准备分公司')]
count_screen20 = screen20.shape[0]

screen21 = screen_read2[(screen_read2[u'所属部门'] == '研发设计中心')]
count_screen21 = screen21.shape[0]

screen22 = screen_read2[(screen_read2[u'所属部门'] == '质检中心')]
count_screen22 = screen22.shape[0]

screen23 = screen_read2[(screen_read2[u'所属部门'] == '后勤服务站')]
count_screen23 = screen23.shape[0]

with pd.ExcelWriter(r'F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx') as writer:
    # merge.to_excel(writer, sheet_name='All', index=False)
    screen4.to_excel(writer, sheet_name='生产科', index=False)
    screen5.to_excel(writer, sheet_name='经营科', index=False)
    screen6.to_excel(writer, sheet_name='党委办公室', index=False)
    screen7.to_excel(writer, sheet_name='安全环保科', index=False)
    screen8.to_excel(writer, sheet_name='技术监督科', index=False)
    screen9.to_excel(writer, sheet_name='企管科', index=False)
    screen10.to_excel(writer, sheet_name='容器制造分公司', index=False)
    screen11.to_excel(writer, sheet_name='修理分公司', index=False)
    screen12.to_excel(writer, sheet_name='光正分公司', index=False)
    screen13.to_excel(writer, sheet_name='自控维护分公司', index=False)
    screen14.to_excel(writer, sheet_name='设备制造分公司', index=False)
    screen15.to_excel(writer, sheet_name='钻采设备分公司', index=False)
    screen16.to_excel(writer, sheet_name='油管检修分公司', index=False)
    screen17.to_excel(writer, sheet_name='防腐分公司', index=False)
    screen18.to_excel(writer, sheet_name='燃烧器项目组', index=False)
    screen19.to_excel(writer, sheet_name='销售部', index=False)
    screen20.to_excel(writer, sheet_name='生产准备分公司', index=False)
    screen21.to_excel(writer, sheet_name='研发设计中心', index=False)
    screen22.to_excel(writer, sheet_name='质检中心', index=False)
    screen23.to_excel(writer, sheet_name='后勤服务站', index=False)


screen_read2_more0 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='生产科')
screen004 = screen_read2_more0[(screen_read2_more0[u'用餐类型'] == '早餐')]
count_screen000 = screen004.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='经营科')
screen005 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen010 = screen005.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='党委办公室')
screen006 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen020 = screen006.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='安全环保科')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen030 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='技术监督科')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen040 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='技术监督科')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen050 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='容器制造分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen060 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='修理分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen070 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='光正分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen080 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='自控维护分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen090 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='设备制造分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0100 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='钻采设备分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0110 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='油管检修分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0120 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类早餐.xlsx', sheet_name='防腐分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0130 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='燃烧器项目组')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0140 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='销售部')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0150 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='生产准备分公司')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0160 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='研发设计中心')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0170 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='质检中心')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0180 = screen00.shape[0]

screen_read2_more1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\部门分类.xlsx', sheet_name='后勤服务站')
screen00 = screen_read2_more1[(screen_read2_more1[u'用餐类型'] == '早餐')]
count_screen0190 = screen00.shape[0]

# date_now = str(datetime.date.today())
# date_tomorrow = (datetime.date.today() + datetime.timedelta(days=+1)).strftime('%Y-%m-%d')

put4 = pd.Series(
    [date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now,
     date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now, date_now],
    index=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20], name='日期')
put5 = pd.Series(
    ['生产科', '经营科', '党委办公室', '安全环保科', '技术监督科', '企管科', '容器制造分公司', '修理分公司', '光正分公司', '自控维护分公司', '设备制造分公司', '钻采设备分公司',
     '油管检修分公司', '防腐分公司', '燃烧器项目组', '销售部', '生产准备分公司', '研发设计中心', '质检中心', '后勤服务站'],
    index=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20], name='所属部门')  # -'+'日期:'+date_now)

put6 = pd.Series([count_screen000, count_screen010, count_screen020, count_screen030, count_screen040, count_screen050,
                  count_screen060, count_screen070, count_screen080, count_screen090, count_screen0100,
                  count_screen0110,
                  count_screen0120, count_screen0130, count_screen0140, count_screen0150, count_screen0160,
                  count_screen0170,
                  count_screen0180, count_screen0190, ],
                 index=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                 name='早餐')

count_2 = pd.DataFrame({put4.name: put4, put5.name: put5, put6.name: put6})

# count_dir1 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')
# count_1 = count_dir1.loc[:, '用餐类型'].value_counts()
# # print(count_1)
# print('用餐类型统计完成')
# count_dir1.to_excel('F:\#!Python_porject\用餐信息数据清洗\总表\A_总表.xlsx', sheet_name='统计用餐类型', )

# count_dir2 = pd.read_excel('F:\#!Python_porject\用餐信息数据清洗\总表\合并.xlsx')
# count_2 = count_dir2.loc[:, '所属部门'].value_counts()
# print(count_2)
# print('所属部门统计完成')
# count_dir2.to_excel('F:\#!Python_porject\用餐信息数据清洗\总表\A_总表.xlsx', sheet_name='统计所属部门')




with pd.ExcelWriter(r'F:\#!Python_porject\用餐信息数据清洗\A早餐_总表.xlsx') as writer:
    merge.to_excel(writer, sheet_name='合并数据', index=False)
    # sort1.to_excel(writer, sheet_name='部门排序', index=False)
    # sort2.to_excel(writer, sheet_name='用餐类型排序', index=False)
    # screen1.to_excel(writer, sheet_name='早餐情况汇总', index=False)
    # screen2.to_excel(writer, sheet_name='午餐情况汇总', index=False)
    # screen3.to_excel(writer, sheet_name='晚餐情况汇总', index=False)
    count_1.to_excel(writer, sheet_name='统计用餐类型', index=False)
    count_2.to_excel(writer, sheet_name='统计所属部门', index=False)
# count_2.to_excel(writer, sheet_name='统计所属部门', index_label='所属部门')

os.chdir(path='F:\#!Python_porject\用餐信息数据清洗')
dir_b = os.getcwd()
print('数据导出完成\n文件路径:' + dir_b)
