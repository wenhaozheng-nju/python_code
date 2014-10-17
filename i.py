#!/usr/bin/env python
# coding=utf-8

import xlrd

data = xlrd.open_workbook('grade.xls')
table = data.sheets()[0]

data2 = xlrd.open_workbook('2011cs.xls')
table2 = data2.sheets()[0]

nrows = table.nrows
ncols = table.ncols

nrows2 = table2.nrows
ncols2 = table2.ncols

dict = {}
for i in range(1,nrows):
    dict[table.cell(i,1).value] = table.cell(i,2).value

#print dict

print '大学物理实验'
count = 0
for j in range(1,nrows2):

    if(dict.get(table2.cell(j,2).value)):
        print dict.get(table2.cell(j,2).value)
    else:
        print '-3'
'''
        if(table2.cell(j,ncols2-2).value > 0):
            print table2.cell(j,ncols2-2).value
        else:
            print '-3'
            count += 1

print count
'''



