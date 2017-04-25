#-*- coding:utf-8 -*-
import re
from StudentsInfo import *
import xlrd
import xdrlib,sys
import xlwt
import datetime
import  time
#得到学生的所有信息
def getStudent(filename):
    students=[]
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    nrows = table.nrows  #行数
    ncols = table.ncols  #列数
    info = table.row_values(0)
    for i in range(ncols):
        if(table.row_values(0)[i]==u'学号'):
            for j in range(1, nrows):
                temp = StudentsInfo()
                temp.id = table.row_values(j)[i]
                students.append(temp)
        if(table.row_values(0)[i]==u'国别'):
            for j in range(1,nrows):
                students[j-1].country=table.row_values(j)[i]
        if (table.row_values(0)[i] == u'英文姓名'):
            for j in range(1, nrows):
                students[j - 1].ename = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'中文姓名'):
            for j in range(1, nrows):
                students[j - 1].cname = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'性别'):
            for j in range(1, nrows):
                students[j - 1].sex = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'学生类别'):
            for j in range(1, nrows):
                students[j - 1].studentcatagory = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'专业'):
            for j in range(1, nrows):
                students[j - 1].major = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'入专业日期'):
            for j in range(1, nrows):
                students[j - 1].joinmajor = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'预计离校日期'):
            for j in range(1, nrows):
                students[j - 1].leaveschool = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'经费办法'):
            for j in range(1, nrows):
                students[j - 1].jingfei = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'系所'):
            for j in range(1, nrows):
                students[j - 1].department = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'入学年份'):
            for j in range(1, nrows):
                students[j - 1].startschool = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'电子邮件'):
            for j in range(1, nrows):
                students[j - 1].email = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'最高学历'):
            for j in range(1, nrows):
                students[j - 1].highesteducation = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'获奖学金情况'):
            for j in range(1, nrows):
                students[j - 1].scholarship = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'学籍状态'):
            for j in range(1, nrows):
                students[j - 1].learnstatus = table.row_values(j)[i]
        if (table.row_values(0)[i] == u'交换学校'):
            for j in range(1, nrows):
                students[j - 1].exchangeschool = table.row_values(j)[i]

    return students

#获取世界排名的学校
def getWord_BestSchool(num):
    bestSchool={}
    data=xlrd.open_workbook(u'US\\US世界排名.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    if num>nrows:
        num=nrows
    for i in range(num):
        if table.row_values(i)[0] not in bestSchool:
            s=table.row_values(i)[0]
            s=s.replace(' ','') #去除空格
            s.lower()           #学校名字全部为小写
            bestSchool[s]=s
    return bestSchool

def majorPercent(filename):
    students = getStudent(filename)
    majorMap = {}
    for key in students:  # 对学校进行统计
        # print key.exchangeschool
        if key.major not in majorMap:
            majorMap[key.major] = 1
        else:
            majorMap[key.major] = majorMap[key.major] + 1
    # 建立excel表头
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u'专业占比')
    sheet1.write(0, 0, u'专业')
    sheet1.write(0, 1, u'人数')
    sheet1.write(0, 2, u'百分比')
    num=1
    for key in majorMap:
        sheet1.write(num,0,key)
        sheet1.write(num,1,majorMap[key])
        d=100*float(majorMap[key])/float(len(students))
        sheet1.write(num,2,"%.2f"%(d))
        num=num+1
    book.save(u'专业百分比.xls')
    return majorMap

def sexPercent(filename):
    students = getStudent(filename)
    sexMap = {}
    for key in students:  # 对学校进行统计
        # print key.exchangeschool
        if key.sex not in sexMap:
            sexMap[key.sex] = 1
        else:
            sexMap[key.sex] = sexMap[key.sex] + 1
    # 建立excel表头
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u'男女占比')
    sheet1.write(0, 0, u'男生人数（人）')
    sheet1.write(0, 1, u'女生人数（人）')
    sheet1.write(0, 2, u'男生占百分比（%）')
    sheet1.write(0, 3, u'女生生占百分比（%）')
    sheet1.write(0, 4, u'合计人数(人)')
    num=1
    for key in sexMap:
        if(key==u'男'):
            sheet1.write(num,0,sexMap[key])
            d=100*float(sexMap[key])/float(len(students))
            sheet1.write(num,2,"%.4f%%"%(d))
        elif(key==u'女'):
            sheet1.write(num, 1, sexMap[key])
            d = 100 * float(sexMap[key]) / float(len(students))
            sheet1.write(num, 3, "%.4f%%" % (d))
    sheet1.write(num, 4, len(students))
    book.save(u'性别百分比.xls')
    return sexMap

def sortSchool(filename,num):
    students = getStudent(filename)
    schoolMap = {}
    for key in students: #对学校进行统计
        # print key.exchangeschool
        if key.exchangeschool not in schoolMap:
            schoolMap[key.exchangeschool] = 1
        else:
            schoolMap[key.exchangeschool] = schoolMap[key.exchangeschool] + 1
    bestSchool=getWord_BestSchool(num)
    #建立excel表头
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u'世界名校的学生')
    sheet1.write(0, 0, u'学号')
    sheet1.write(0, 1, u'国别')
    sheet1.write(0, 2, u'英文姓名')
    sheet1.write(0, 3, u'中文姓名')
    sheet1.write(0, 4, u'性别')
    sheet1.write(0, 5, u'学生类别')
    sheet1.write(0, 6, u'专业')
    sheet1.write(0, 7, u'入专业日期')
    sheet1.write(0, 8, u'预计离校日期')
    sheet1.write(0, 9, u'经费办法')
    sheet1.write(0, 10, u'系所')
    sheet1.write(0, 11, u'入学年份')
    sheet1.write(0, 12, u'电子邮件')
    sheet1.write(0, 13, u'最高学历')
    sheet1.write(0, 14, u'获奖学金情况')
    sheet1.write(0, 15, u'学籍状态')
    sheet1.write(0, 16, u'交换学校')
    num=0
    for key in students:
        # "[清华大学]Tsinghua university"
        #去除[清华大学]
        com = re.compile('\[.*\]')
        s =key.exchangeschool
        ma = com.match(s)
        t = s[s.find(']') + 1:]
        t = t.replace(' ', '')  # 去除空格
        t.lower()  # 学校名字全部为小写
        if t in bestSchool: # 如果在世界排名靠前的名校则将信息写入 excel
            num=num+1
            sheet1.write(num, 0, key.id)
            sheet1.write(num, 1, key.country)
            sheet1.write(num, 2, key.ename)
            sheet1.write(num, 3, key.cname)
            sheet1.write(num, 4, key.sex)
            sheet1.write(num, 5, key.studentcatagory)
            sheet1.write(num, 6, key.major)
            sheet1.write(num, 7, key.joinmajor)
            sheet1.write(num, 8, key.leaveschool)
            sheet1.write(num, 9, key.jingfei)
            sheet1.write(num, 10, key.department)
            sheet1.write(num, 11, key.startschool)
            sheet1.write(num, 12, key.email)
            sheet1.write(num, 13, key.highesteducation)
            sheet1.write(num, 14, key.scholarship)
            sheet1.write(num, 15, key.learnstatus)
            sheet1.write(num, 16, key.exchangeschool)


    book.save(u'留学生世界排名结果.xls')

if __name__ == '__main__':
    sexPercent('test.xls')