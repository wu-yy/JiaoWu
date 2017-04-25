# -*- coding: utf-8 -*-

from Tkinter import *
from matplotlib import pyplot as plt
import matplotlib as mpl

mpl.rcParams['font.family'] = 'sans-serif'
mpl.rcParams['font.sans-serif'] = [u'SimHei']

import Tkinter

from tkMessageBox import *

import sys

import Pmw

import time

import os
from sortSchool import  *

def plotPictures(labels,sizes,filename):
    # 调节图形大小，宽，高
    plt.figure(figsize=(6, 9))
    # 定义饼状图的标签，标签是列表
    # 每个标签占多大，会自动去算百分比
    colors = ['red', 'yellow','green','blue','gray','lightgreen','olive','darkgreen']

    # 将某部分爆炸出来， 使用括号，将第一块分割出来，数值的大小是分割出来的与其他两块的间隙
    #explode = (0.05, 0)

    patches, l_text, p_text = plt.pie(sizes, explode=None, labels=labels, colors=colors,
                                      labeldistance=1.1, autopct='%3.1f%%', shadow=False,
                                      startangle=90, pctdistance=0.9)

    # labeldistance，文本的位置离远点有多远，1.1指1.1倍半径的位置
    # autopct，圆里面的文本格式，%3.1f%%表示小数有三位，整数有一位的浮点数
    # shadow，饼是否有阴影
    # startangle，起始角度，0，表示从0开始逆时针转，为第一块。一般选择从90度开始比较好看
    # pctdistance，百分比的text离圆心的距离
    # patches, l_texts, p_texts，为了得到饼图的返回值，p_texts饼图内部文本的，l_texts饼图外label的文本

    # 改变文本的大小
    # 方法是把每一个text遍历。调用set_size方法设置它的属性
    for t in l_text:
        t.set_size = (30)
    for t in p_text:
        t.set_size = (20)
    # 设置x，y轴刻度一致，这样饼图才能是圆的
    plt.axis('equal')
    #plt.legend()
    plt.savefig(filename)
    plt.show()
class GUIFrame(Frame):
    """ Demonstrate Entrys and Envent Building"""


    first = Init_FileCount = "test.xls"

    Init_FileCount_Postfix = ".xls"

    # File_Path = os.getcwd()

    File_Path = ".//"

    print "File_Path=", File_Path

    def showContents(self, event):

        print "showContents"
        theName = event.widget.winfo_name()

        theContents = event.widget.get()

        showinfo("Message", theName + ":" + theContents)

    def showSortSchool(self):
        temp=self.text1.get()
        num=self.text2.get()
        sortSchool(temp,num)

    def showSexPercent(self):
        sexMap=sexPercent(self.text1.get())
        temp=u'男生人数（人） 女生人数(人) 男生占百分比(%) 女生占百分比（%） 合计人数（人）\n'
        self.historyText.insert(INSERT,temp)

        self.historyText.insert(INSERT, " %d " % (sexMap[u'男']))
        d1 = 100 * float(sexMap[u'男']) / float(sexMap[u'男']+sexMap[u'女'])
        self.historyText.insert(INSERT," %.4f%%" % (d1))
        self.historyText.insert(INSERT, " %d " % (sexMap[u'女']))
        d2 = 100 * float(sexMap[u'女']) / float(sexMap[u'男']+sexMap[u'女'])
        self.historyText.insert(INSERT, " %.4f%% " % (d2))
        self.historyText.insert(INSERT, " %d \n" % (sexMap[u'女']+sexMap[u'男']))
        labels=[u'男',u'女']
        sizes=[sexMap[u'男'],sexMap[u'女']]
        plotPictures(labels=labels,sizes=sizes,filename=u'性别百分比.png')

    def showMajorpercent(self):

        majorMap=majorPercent(self.text1.get())
        temp=u'专业 人数  百分比\n'
        self.historyText.insert(INSERT, temp)
        stu=0
        for key in majorMap:
            stu=stu+majorMap[key]

        labels=[]
        sizes=[]
        for key in majorMap:
            labels.append(key)
            sizes.append(majorMap[key])
            d = 100*float(majorMap[key]) / float(stu)
            self.historyText.insert(INSERT, key)
            self.historyText.insert(INSERT, ' '+str(majorMap[key])+' ')
            self.historyText.insert(INSERT,"%.4f%%" % (d)+'\n')

        plotPictures(labels=labels, sizes=sizes, filename=u'专业百分比.png')

    def showTest(self):
        print self.IdCheck_choose.get()
        #self.historyText.insert(INSERT,'hello\n')

    def __init__(self, parent):

        Frame.__init__(self)
        self.pack(expand=YES, fill=BOTH)
        self.master.title("清华大学留学生数据处理V1.0发布")
        self.master.geometry("800x600-20+20")  # width X length
        self.master.resizable(width=False, height=False)
        # Frame1
        self.frame1 = Frame(self)
        self.frame1.pack(pady=5)  # 垂直间距
        # 文件名输入
        self.label1 = Label(self.frame1, font="Tahoma 10", text="要处理的文件名:")
        # self.spacelabel = Label(self.frame1, width =15)
        self.text1 = Entry(self.frame1, name='text1', width=30)
        self.text1.insert(INSERT, self.Init_FileCount)
        # 记录数
        self.label2 = Label(self.frame1, font="Tahoma 10", text="世界排名数（1-100）:", width=20)
        self.text2 = Entry(self.frame1)
        self.text2.insert(INSERT, "100")
        self.text2.bind("<Return>", self.showContents)
        #button 生成按钮
        self.button2=Button(self.frame1,text='生成',font='Toahoma 10',command=self.showSortSchool)
        self.button2.bind("<Enter>", None)  # 鼠标事件:进入
        self.button2.bind("<Leave>", None)  # 鼠标事件：离开
        self.text1.bind("<Return>", self.showContents)
        self.label1.pack(side=LEFT, padx=5)
        self.text1.pack(side=LEFT, padx=2)
        self.label2.pack(side=LEFT, padx=5)
        self.text2.pack(side=LEFT, padx=2)
        self.button2.pack(side=LEFT, padx=5)

        # Frame2
        # 开始文件数
        self.frame2 = Frame(self)
        self.frame2.pack(pady=10)

        self.IdLab = Label(self.frame2, text="学生ID:", font="Toahoma 10")
        self.IdLab.pack(side=LEFT, padx=3)
        self.IdCheck_choose = BooleanVar()
        print self.IdCheck_choose
        self.IdCheck = Checkbutton(self.frame2, variable=self.IdCheck_choose, font="Toahoma 10", command=self.showTest)
        self.IdCheck.pack(side=LEFT, padx=3)

        self.telType = Label(self.frame2, text="学生类型:", font="Toahoma 10")
        self.telType_choose = BooleanVar()
        self.telCheck = Checkbutton(self.frame2, variable=self.telType_choose, font="Toahoma 10", command=None)
        self.telType.pack(side=LEFT, padx=3)
        self.telCheck.pack(side=LEFT, padx=3)

        # Frame3

        # 展示专业百分比

        self.frame3 = Frame(self)
        self.frame3.pack(pady=20)
        self.spacelabel3 = Label(self.frame3, width=30)
        self.label6 = Label(self.frame3, font="Toahoma 10", text="展示专业百分比:")
        self.courseButton = Button(self.frame3, text='生成', font='Toahoma 10', command=self.showMajorpercent)
        self.label7 = Label(self.frame3, font="Toahoma 10", text="展示男女比:")
        self.sexButton = Button(self.frame3, text='生成', font='Toahoma 10', command=self.showSexPercent)
        self.label6.pack(side=LEFT, padx=5)
        self.courseButton.pack(side=LEFT,padx=5)
        self.label7.pack(side=LEFT,padx=5)
        self.sexButton.pack(side=LEFT,padx=5)


        # frame4

        # 消息内容
        self.frame4 = Frame(self)
        self.frame4.pack(pady=10)
        self.spacelabel4 = Label(self.frame4, width=5)
        self.contentLab = Label(self.frame4, text="内容:", font="Toahoma 10")
        self.contentLab.pack(side=LEFT, padx=3)
        self.historyText = Pmw.HistoryText(self.frame4,

                                           text_wrap='none',

                                           text_width=77,

                                           text_height=20,

                                           )

        self.historyText.pack(side=LEFT)
        self.historyText.component('text').focus()
        self.spacelabel4.pack(side=LEFT, padx=3)

        '''
        # frame7

        # 自动内容识别

        self.frame7 = Frame(self)
        self.frame7.pack(pady=10)
        self.items = (i for i in range(10))
        self.autoLab = Label(self.frame7, text="自动生成内容基础条数:", font="Toahoma 10")
        self.autoLab.pack(side=LEFT, padx=3)
        self.dropdown = Pmw.ComboBox(self.frame7,

                                     #           labelpos = 'nw',

                                     #          selectioncommand = self.changeColour,

                                     scrolledlist_items=self.items,

                                     entry_width=7

                                     )

        self.dropdown.pack(side=LEFT, padx=5)
        '''
        # frame9
        self.frame9 = Frame(self)

        self.frame9.pack(pady=10)

        self.messageBar = Pmw.MessageBar(self.frame9,

                                         entry_width=40,

                                         entry_relief='groove',

                                         labelpos='w',

                                         label_text='Status:')

        self.messageBar.pack(fill='x', expand=1, padx=10, pady=5)

        # frame8

        # 版权信息

        self.frame8 = Frame(self)

        self.frame8.pack(pady=10)

        self.versionLab = Label(self.frame8, text="Copyright @2017 Tsinghua TestWork SoftWare .", font="Toahoma 9")

        self.versionLab.pack(side=LEFT, padx=3)

def main():
    import Tkinter

    root = Tkinter.Tk()

    Pmw.initialise(root)

    widget = GUIFrame(root)

    root.mainloop()

if __name__ == "__main__":
    main()