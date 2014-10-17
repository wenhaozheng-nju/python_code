#!/usr/bin/env python
# coding=utf-8

import xlrd



data = xlrd.open_workbook('2011cs.xls')
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols
credit = (5,5,4,3,4,4,4,4,4,2,4,4,2,4,2,4,4,4,4,3,3,2,3,3,3,3,2)
#wb = copy(data)
#ws = wb.get_sheet(0)
print 'GPA'
count = 1
my_count = 0
for i in range(1,nrows):
    sum = 0
    credit_sum = 0
    flag = 1
    for j in range(len(credit)):
        value = table.cell(i,j+3).value
        if(value > 0):
            sum += value * credit[j]
            credit_sum += credit[j]

        if(value <= 0 and j < len(credit)):
           flag = 0
    
#    if not flag:
#        print table.cell(i,2).value
#       my_count += 1
#    if(flag):
#        print count,table.cell(i,2).value,
    print (sum + 0.0) / (credit_sum * 20)
#print my_count
    #count += 1
#    table.put_cell(i,ncols-2,2,(sum+0.0)/credit_sum,0)
#    table.cell(i,ncols-2).value = (sum + 0.0)/credit_sum
#    ws.write(i,ncols-2,(sum + 0.0)/credit_sum)
