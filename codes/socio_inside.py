import xlrd
import xlwt
import pandas as pd
import matplotlib.pyplot as plt

# get significant features
workbook = xlrd.open_workbook('ACS/ACS_13_5YR_DP02_with_ann.xlsx')
sheet = workbook.sheet_by_index(0)
n = sheet.nrows - 1
quarter = 100
indexes = list()
factors = list()
conditions = list()
col_names = sheet.row_values(0) #index
for i in range(148):
    j = 6+4*i
    if j not in [62,66,154,158,162,166,170,282,286,290,294,298,302,306,310]:
        indexes.append(j)
        factors.append(list())
#初始化阈值
index_num = len(indexes)
for i in range(index_num):
    col = sheet.col_values(indexes[i])[1:]
    while '-' in col:
        col[col.index('-')] = 0
    while '(X)' in col:
        col[col.index('(X)')] = 0
    col.sort()
    conditions.append([col[3 * quarter], col[2 * quarter], col[quarter]])

for i in range(index_num):
    index = indexes[i]
    col = sheet.col_values(indexes[i])[1:]
    while '-' in col:
        col[col.index('-')] = 0
    while '(X)' in col:
        col[col.index('(X)')] = 0
    for num in col:
        if num >= conditions[i][0]:
            factors[i].append(col_names[index] + "_1")
        elif num < conditions[i][2]:
            factors[i].append(col_names[index] + "_4")
        elif conditions[i][0] > num >= conditions[i][1]:
            factors[i].append(col_names[index] + "_2")
        else:
            factors[i].append(col_names[index] + "_3")

# workbook = xlwt.Workbook(encoding ='utf-8')
# name = "2013(1)"
# worksheet = workbook.add_sheet(name)
# worksheet.write(0,0,"drug")
# for i in range(30):
#     worksheet.write(0, i + 1, col_names[indexes[i]])
# for i in range(n):
#     if i <= quarter:
#         worksheet.write(i+1,0,"high")
#     elif i > 3*quarter:
#         worksheet.write(i+1,0,"low")
#     else:
#         worksheet.write(i+1,0,"mid")
#     for j in range(30):
#         index = indexes[j]
#         worksheet.write(i+1,j+1,factors[j][i])
# workbook.save(name+'.xls')

# workbook = xlwt.Workbook(encoding ='utf-8')
# name = "2013(2)"
# worksheet = workbook.add_sheet(name)
# worksheet.write(0,0,"drug")
# for i in range(30):
#     worksheet.write(0, i + 1, col_names[indexes[30+i]])
# for i in range(n):
#     if i <= quarter:
#         worksheet.write(i+1,0,"high")
#     elif i > 3*quarter:
#         worksheet.write(i+1,0,"low")
#     else:
#         worksheet.write(i+1,0,"mid")
#     for j in range(30):
#         index = indexes[j]
#         worksheet.write(i+1,j+1,factors[30+j][i])
# workbook.save(name+'.xls')

# workbook = xlwt.Workbook(encoding ='utf-8')
# name = "2013(3)"
# worksheet = workbook.add_sheet(name)
# worksheet.write(0,0,"drug")
# #50+i,50+j
# #index_num-i-1,index_num-j-1
# for i in range(30):
#     worksheet.write(0, i + 1, col_names[indexes[60+i]])
# for i in range(n):
#     if i <= quarter:
#         worksheet.write(i+1,0,"high")
#     elif i > 3*quarter:
#         worksheet.write(i+1,0,"low")
#     else:
#         worksheet.write(i+1,0,"mid")
#     for j in range(30):
#         index = indexes[j]
#         worksheet.write(i+1,j+1,factors[60+j][i])
# workbook.save(name+'.xls')

# workbook = xlwt.Workbook(encoding ='utf-8')
# name = "2013(4)"
# worksheet = workbook.add_sheet(name)
# worksheet.write(0,0,"drug")
# #50+i,50+j
# #index_num-i-1,index_num-j-1
# for i in range(30):
#     worksheet.write(0, i + 1, col_names[indexes[index_num-i-1]])
# for i in range(n):
#     if i <= quarter:
#         worksheet.write(i+1,0,"high")
#     elif i > 3*quarter:
#         worksheet.write(i+1,0,"low")
#     else:
#         worksheet.write(i+1,0,"mid")
#     for j in range(30):
#         index = indexes[j]
#         worksheet.write(i+1,j+1,factors[index_num-j-1][i])
# workbook.save(name+'.xls')

data_xls = pd.read_excel(name+'.xls', sheet_name=name)
data_xls.to_csv(name+'.csv', encoding='utf-8')

# workbook = xlrd.open_workbook('ACS/ACS_16_5YR_DP02_with_ann.xlsx')
# sheet = workbook.sheet_by_index(0)
# n = sheet.nrows - 1
# quarter = 100
# indexes = [34,118,242]
# index_num = len(indexes)
# factors = list()
# conditions = list()
# col_names = sheet.row_values(0) #index
# for i in range(index_num):
#     factors.append(list())
# #初始化阈值
# index_num = len(indexes)
# for i in range(index_num):
#     col = sheet.col_values(indexes[i])[1:]
#     while '-' in col:
#         col[col.index('-')] = 0
#     col.sort()
#     conditions.append([col[3 * quarter], col[2 * quarter], col[quarter]])
#
# for i in range(index_num):
#     index = indexes[i]
#     col = sheet.col_values(indexes[i])[1:]
#     while '-' in col:
#         col[col.index('-')] = 0
#     for num in col:
#         if num >= conditions[i][0]:
#             factors[i].append(col_names[index] + "_1")
#         elif num < conditions[i][2]:
#             factors[i].append(col_names[index] + "_4")
#         elif conditions[i][0] > num >= conditions[i][1]:
#             factors[i].append(col_names[index] + "_2")
#         else:
#             factors[i].append(col_names[index] + "_3")
#
# workbook = xlwt.Workbook(encoding ='utf-8')
# worksheet = workbook.add_sheet('2016')
# worksheet.write(0,0,"drug")
# for i in range(index_num):
#     worksheet.write(0, i + 1, col_names[indexes[i]])
# for i in range(n):
#     if i <= quarter:
#         worksheet.write(i+1,0,"high")
#     elif i > 3*quarter:
#         worksheet.write(i+1,0,"low")
#     else:
#         worksheet.write(i+1,0,"mid")
#     for j in range(index_num):
#         index = indexes[j]
#         worksheet.write(i+1,j+1,factors[j][i])
# workbook.save('2016.xls')
#
# data_xls = pd.read_excel('2016.xls', sheet_name="2016")
# data_xls.to_csv('2016.csv', encoding='utf-8')