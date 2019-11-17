import xlrd
import xlwt
import pandas as pd
import matplotlib.pyplot as plt


def averagenum(num):
    nsum = 0
    length = len(num)
    for i in range(length):
        nsum += num[i]
    if length == 0:
        length = 1
    return nsum / length

#[[],[]]
def get_growth(factors):
    for i in range(len(factors)):
        for j in range(len(factors[i]) - 1):
            factors[i][j] = factors[i][j+1] - factors[i][j]
        factors[i].pop(len(factors[i]) - 1)
    return factors

#[[],[],[]]
def offset(l1):
    l3 = list()
    for i in range(len(l1[0])):
        l3.append(l1[0][i]+l1[1][i]-l1[2][i])
    return l3

# get mean of factors in each year in Ohio
# indexes = [18,26,34,38,270]
# index_num = len(indexes)
# factors = list()
# for i in range(index_num):
#     factors.append(list())
# for year in range(10, 17):
#     workbook2 = xlrd.open_workbook('xlsx/ACS_' + str(year) + '_5YR_DP02_with_ann.xlsx')
#     sheet = workbook2.sheet_by_index(0)
#     for i in range(index_num):
#         # col = sheet.col_values(indexes[i] - 1)[2:122]
#         # col = sheet.col_values(indexes[i] - 1)[122:210]
#         # col = sheet.col_values(indexes[i] - 1)[210:277]
#         col = sheet.col_values(indexes[i] - 1)[277:411]
#         # col = sheet.col_values(indexes[i] - 1)[411:]
#         factors[i].append(averagenum(col))
# print(factors)
# workbook2 = xlrd.open_workbook('xlsx/ACS_10_5YR_DP02_with_ann.xlsx')
# sheet = workbook2.sheet_by_index(0)
# col_names = sheet.row_values(1)

marriage = [0.39111,0.36940,0.35625,0.42975]

ax = plt.subplot(111)
x = range(4)
ax.set_ylabel("$offset_{G_{p,m}}$")
# ax.set_ylabel("$offset_{G_{p,m},G_{p,b}}$")
# ax.set_ylabel("$offset_{G_{p,nw},G_{p,nh},G_{p,nhc}}$")
# plt.plot(x,states,label="$offset_{G_{p,m},G_{p,b}}$")
plt.plot(x,marriage,label="$G_{p,m}$")
plt.xticks(x,("WV","PA","KY","OH"))
# plt.legend()
plt.show()