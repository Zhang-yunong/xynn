import xlwt
import xlrd
import math
import random
import numpy as np
wbk = xlwt.Workbook()
sheet_ideal = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)  #创建列表保存随机志愿
people_num = 395   #总人数
items = ['CK', 'GD', 'DZ', 'SW']
#CK：测控技术与仪器 GD：光电信息科学与工程 DZ：电子科学与技术 SW：生物医学工程
#people_ideal = np.zeros((people_num,5), dtype=np.str)
people_ideal = [[0] * 5 for _ in range(people_num)] #创建志愿列表
CK_people_num = round(people_num * 0.4)    #158 各专业名额，四舍五入
GD_people_num = round(people_num * 0.25)   #99
DZ_people_num = round(people_num * 0.18)   #71
SW_people_num = round(people_num * 0.17)   #67
CK_people = [0] * CK_people_num #创建选志愿名单
GD_people = [0] * GD_people_num
DZ_people = [0] * DZ_people_num
SW_people = [0] * SW_people_num
name = ['排名','第一志愿','第二志愿','第三志愿','第四志愿']
name2 = ['CK','GD ','DZ','SW']
for i in range(5):
    sheet_ideal.write(0,i,name[i])    #写表头
for i in range(people_num):
    random.shuffle(items)         #生成随机志愿顺序
    people_ideal[i][0] = i+1      #生成排名
    sheet_ideal.write(i+1,0,people_ideal[i][0])  #排名写入列表
    for j in range(4):
        people_ideal[i][j+1] = items[j]        #生成志愿
        sheet_ideal.write(i+1,j+1,people_ideal[i][j+1])  #志愿写入列表

wbk.save('ideal.xls') #保存列表

wbk = xlwt.Workbook()
sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True) #创建列表保存选志愿结果

data = xlrd.open_workbook('ideal.xls') #打开一个表格
table = data.sheets()[0] # 打开第一张表
nrows = table.nrows      # 获取表的行数
for i in range(people_num):
    for j in range(5):
        people_ideal[i][j] = table.row_values(i+1)[j]   #读取排名和志愿
for i in range(4):
    sheet.write(0,i,name2[i])   #写表头
        
def select(people_num):
    CK_num = 0 #已选人数
    GD_num = 0
    DZ_num = 0
    SW_num = 0
    for i in range(people_num):
        j = 1  #志愿序号初始化
        while j < 5 :
            if people_ideal[i][j] == 'CK' and CK_num < CK_people_num:
                CK_people[CK_num] = people_ideal[i][0]
                sheet.write(CK_num+1, 0, CK_people[CK_num])
                CK_num += 1
                j += 4
            elif people_ideal[i][j] == 'GD' and GD_num < GD_people_num:
                GD_people[GD_num] = people_ideal[i][0]
                sheet.write(GD_num+1, 1, GD_people[GD_num])
                GD_num += 1
                j += 4
            elif people_ideal[i][j] == 'DZ' and DZ_num < DZ_people_num:
                DZ_people[DZ_num] = people_ideal[i][0]
                sheet.write(DZ_num+1, 2, DZ_people[DZ_num])
                DZ_num += 1
                j += 4
            elif people_ideal[i][j] == 'SW' and SW_num < SW_people_num:
                SW_people[SW_num] = people_ideal[i][0]
                sheet.write(SW_num+1, 3, SW_people[SW_num])
                SW_num += 1
                j += 4
            else :
                j += 1

select(people_num)

wbk.save('result.xls') #保存列表



