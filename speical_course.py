#!/usr/bin/env python
# coding=utf-8

import xlrd
import xlwt

except_name = ['105220001','111220009','111220070','111220158','111220151','115220001','111220102','111220101','111220025','101220161','091220095','101220049','101220082','101220107']
speical_name = ['111220033','111220112','111220005','111100027','111220018','111220153','111170028','111220095','111220196','111220107','111220073','111250148','111250105','111220172','111220127','111220075','111220171','111220081','111242059','111242033']
math_physical = ['120011','120010A','000111','000121','000141']

num2name = {}

def input_grade(name,dict):
    data = xlrd.open_workbook(name)
    table = data.sheets()[0]
    nrows = table.nrows
    for i in range(1, nrows): 
        if cmp(table.cell(i, 4).value, u'注销') != 0 and table.cell(i,5).value and float(table.cell(i,5).value) >= 60:
            st_num = table.cell(i, 0).value
            if not dict.get(st_num):
                dict[st_num] = {}
                num2name[st_num] = table.cell(i,1).value
            course_name = table.cell(i, 2).value
            if table.cell(i,6).value:
                str = table.cell(i,6).value
                str_array = str.split()
                index = 0
                for k in str_array[0]:
                    if k.isdigit() or k.isalpha():
                        index += 1
                    else:
                        break;
                str_array[1] = str_array[0][index+1]
                str_array[0] = str_array[0][0:index]
#                print str_array[0],str_array[1],type(dict[st_num])
                if cmp(str_array[0][0:2],'22') == 0 or str_array[0] in math_physical: #此处如是计算普通班，需加入数学物理类课程 
                    if dict.get(str_array[0]):
                        dict[st_num][str_array[0]][1] = table.cell(i,5).value
                    else:
                        dict[st_num][str_array[0]] = []
                        dict[st_num][str_array[0]].append(str_array[1])
                        dict[st_num][str_array[0]].append(table.cell(i,5).value)
            elif cmp(course_name[0:2],'22') == 0 or course_name in math_physical:
                if dict.get(course_name):
                    dict[st_num][course_name][1] = table.cell(i, 5).value
                else:
                    dict[st_num][course_name] = [table.cell(i, 3).value, table.cell(i, 5).value]

if __name__ == '__main__':
    dict = {} 
    grade_file = xlwt.Workbook()
    grade_table = grade_file.add_sheet('sheet0')
    course_file = xlwt.Workbook()
    course_table = course_file.add_sheet('sheet0')
    grade_table.write(0,0,u'学号')
    grade_table.write(0,1,u'姓名')
    grade_table.write(0,2,'GPA')
    course_table.write(0,0,u'学号')
    course_table.write(0,1,u'课程编号')
    course_table.write(0,2,u'学分')
    course_table.write(0,3,u'总评')
#    input_grade('大零上成绩.xls',dict)
#    input_grade('大零下成绩.xls',dict)
#    input_grade('大一上成绩.xls',dict)
#    input_grade('大一下成绩.xls',dict)
#    input_grade('大二上成绩.xls',dict)
#    input_grade('大二下成绩.xls',dict)
    input_grade('大三上成绩.xls',dict)
    input_grade('大三下成绩.xls',dict)
#    input_grade('11.xls',dict)
#    input_grade('21.xls',dict)
#    input_grade('31.xls',dict)
#    input_grade('41.xls',dict)
#    input_grade('12.xls',dict)
#    input_grade('22.xls',dict)
#    input_grade('32.xls',dict)
#    input_grade('42.xls',dict)
    grade_flag = 1
    course_flag = 1
    for i in dict:
        sum = 0.0
        credit = 0.0
        if i in except_name or i in speical_name:
            continue
        for j in dict[i]:
#            print i,j,dict[i][j][0],dict[i][j][1]
            course_table.write(course_flag,0,i.decode('utf-8'))
            course_table.write(course_flag,1,j.decode('utf-8'))
            course_table.write(course_flag,2,dict[i][j][0])
            course_table.write(course_flag,3,dict[i][j][1])
            course_flag += 1
#        print dict[i][j][0]," ",dict[i][j][1]
            credit += float(dict[i][j][0])
            sum += (float(dict[i][j][1]) * float(dict[i][j][0]))            
#        print "%s %s %f %f %f" % (i,num2name[i],(sum+0.0)/(credit*20),credit,sum)
        grade_table.write(grade_flag,0,i)
        grade_table.write(grade_flag,1,num2name[i])
        if credit == 0:
            grade_table.write(grade_flag,2,0)
        else:
            grade_table.write(grade_flag,2,(sum+0.0)/(credit*20))
        grade_flag += 1
        
    grade_file.save('../grade/大二专业学分绩.xls')
    course_file.save('../grade/大二专业学分绩计算课程.xls')
