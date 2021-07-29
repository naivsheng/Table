# -*- coding: UTF-8 -*-
# __author__ = 'Yingyu Wang'

from tkinter import *
#from tkinter import ttk
import tkinter as tk
from tkinter.font import Font
from tkinter.ttk import *
from tkinter.messagebox import *
import os
import pandas as pd
from TableReader import TableReader
import time
from datetime import datetime
from tkinter import scrolledtext

class Application_ui(Frame):
    #这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    # S(Soll): 应订     B(Bestellen):订货   K(Kontrollen):收货   R(Rechnung):发票    E(Eingeben):录入   W(Beschwerbe):已投诉  
    # y:点货后需投诉    i(in ordnung)：不需要投诉    o：发送订单     Z:翟到货拍照
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('订货录入辅助')
        self.master.geometry('1000x800')
        # global filePath
        filePath = os.getcwd()
        # path = 'H:\\py\\test\\goasia'
        # os.chdir(path)
        # file_list = os.listdir(path)
        self.FL = ['']
        df = pd.read_excel('Filialen.xlsx',header = 0)
        self.FL.extend(df.iloc[:,0])
        self.FL = {x: self.FL[x] for x in range(len(self.FL))}
        self.FLd = {self.FL[x]:x-1 for x in range(len(self.FL))}
        self.LF = ['']
        df = pd.read_excel('Lieferant.xlsx',header = 0)
        self.LF.extend(df.iloc[:,0])
        self.LFd = {self.LF[x]: x for x in range(len(self.LF))}
        self.createWidgets()
    
    def check(self): # TODO 模糊查询
        self.L = []
        e = self.c.get()
        print(e)
        if not e: self.L = list(self.FL.values())
        for i in self.FLd:
            if e in i:
                self.L.append(i)
        if not self.L: self.L = list(self.FL.values())
        # if L: return(L)
        # else: return(self.FL.values())

    def createWidgets(self):
        '''
            # 创建主窗口：订货、收货、录入、查看
            状态栏公用
            文件打开、更改时间记录/刷新提醒
            TODO：订货总览表建立
            gmailAPI读取：确认发票信息（从pdf附件读取供货商、收货门店信息）
        '''
        self.top = self.winfo_toplevel()    # 主窗口 
        self.style = Style()
        l0 = tk.Label(self.top,text='当前状态',width=8)
        l0.place(x=10,y=10)
        self.var = tk.StringVar()    # 将label标签的内容设置为字符类型，接收b1_clicked函数的传出内容, 显示当前操作状态
        l = tk.Label(self.top, textvariable=self.var, bg='white', fg='black', font=('Arial', 12), width=5, height=1)
        l.place(x=80,y=10)
        l1 = tk.Label(self.top,text='KW=',width=5,height=1)
        l1.place(x=140,y=10)
        self.week = tk.StringVar()
        woche = time.strftime("%W")
        self.week.set(woche)
        entry_week = tk.Entry(self.top,textvariable=self.week,width=5)
        entry_week.place(x=180,y=10)
        '''
        self.c = tk.StringVar() # 用于模糊查询
        entry = tk.Entry(self.top,textvariable=self.c,width=5)
        entry.place(x=550,y=10)
        entry_button = tk.Button(self.top,text='查询', font=('Arial', 12), width=10, height=1, command=self.check)
        entry_button.place(x=650,y=10)
        '''
        # 建立下拉选框，选择分店
        l2 = tk.Label(self.top,text='分店',width=5,height=1)
        l2.place(x=240,y=10)
        valueF = StringVar()
        self.cbxf = tk.ttk.Combobox(self.top, width = 8, height = 20, textvariable = valueF,state='readonly') #, postcommand = self.show_select)
        # self.cbxf["value"] = list(self.FL.values())
        self.L = list(self.FL.values())
        self.cbxf['value'] = self.L
        self.cbxf.place(x=280,y=10)
        l6 = tk.Label(self.top,text='供应商',width=6,height=1)
        l6.place(x=400,y=10)
        valueL = StringVar()
        self.cbxl = tk.ttk.Combobox(self.top,width = 10, height = 20, textvariable = valueL,state='readonly')
        self.cbxl["value"] = self.LF
        self.cbxl.place(x=450,y=10)
        #self.on_hit = False
        
        '''
        # TODO: 模糊查找 Entry + Combobox
        E = tk.StringVar() # 用于存入模糊查找的输入数据
        E.set('')
        valueF = StringVar()
        e1 = tk.Entry(self.top,textvariable=E,width=6)
        # self.a = str(E.get())
        L = e1.bind('<KeyRelease>', self.check(E.get()))
        self.cbxf = tk.ttk.Combobox(self.top, width = 6, height = 1, textvariable = valueF,state='readonly') #, postcommand = self.show_select)
        L = self.check(E.get())
        self.cbxf["value"] = L
        e1.place(x=280,y=10)
        self.cbxf.place(x=300,y=10)
        '''
        
        self.TabStrip1 = Notebook(self.top)
        # 标签页1 订货信息
        self.TabStrip1.place(relx=0.02,rely=0.05,relwidth=0.95,relheight=0.95)
        self.TabStrip1__Tab1 = Frame(self.TabStrip1)
        # 点击 查看 按钮，显示 应订、未订
        b1 = tk.Button(self.TabStrip1__Tab1, text='查看', font=('Arial', 12), width=10, height=1, command=self.bestell_checked)
        b1.place(x=750,y=25)
        # 点击 确认 按钮，将 已订 部分存入xls
        b2 = tk.Button(self.TabStrip1__Tab1, text='确认', font=('Arial', 12), width=10, height=1, command=self.bestell_confirm)
        b2.place(x=750,y=65)
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.TabStrip1__Tab1,text='打开文件时间',width=12,font=('Arial',13))
        tl1.place(x=750,y=300)
        self.b_open_time = tk.StringVar()    
        tl2 = tk.Label(self.TabStrip1__Tab1, textvariable=self.b_open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=750,y=350)
        tl3 = tk.Label(self.TabStrip1__Tab1,text='当前文件时间',width=12,font=('Arial',13))
        tl3.place(x=750,y=400)
        self.b_file_time = tk.StringVar()    
        tl4 = tk.Label(self.TabStrip1__Tab1, textvariable=self.b_file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=750,y=450)
        l3 = tk.Label(self.TabStrip1__Tab1,text='应订',width=5,height=1)
        l3.place(x=60,y=10)
        sb = Scrollbar(self.TabStrip1__Tab1) # 给列表增加滚动条，以防过多数据
        # self.blist1 = Listbox(self.TabStrip1__Tab1,width=20,height=45)
        # self.blist1.place(x=10,y=30)
        self.blist1 = Listbox(self.TabStrip1__Tab1,width=20,height=45,yscrollcommand=sb.set)
        self.blist1.place(x=10,y=30)
        sb.config(command=self.blist1.yview)
        l4 = tk.Label(self.TabStrip1__Tab1,text='未订',width=5,height=1)
        l4.place(x=300,y=10)
        self.blist2 = Listbox(self.TabStrip1__Tab1,width=20,height=45,selectmode = MULTIPLE)
        self.blist2.place(x = 250,y = 30)
        #self.list2.bind('<Double-Button-1>',self.show_select)
        l5 = tk.Label(self.TabStrip1__Tab1,text='本次已订',width=8,height=1)
        l5.place(x=540,y=10)
        self.blist3 = Listbox(self.TabStrip1__Tab1,width=20,height=45,selectmode = MULTIPLE)
        self.blist3.place(x = 500,y = 30)
        #self.list3.insert('')
        #self.list3.bind('<Double-Button-1>',self.delect_select)
        b3 = tk.Button(self.TabStrip1__Tab1, text='订货', font=('Arial', 12), width=5, height=1, command=self.b_show_select)
        b3.place(x=420,y=280)
        b4 = tk.Button(self.TabStrip1__Tab1, text='取消', font=('Arial', 12), width=5, height=1, command=self.b_delect_select)
        b4.place(x=420,y=320)
        info = '应订 显示由之前的数据计算出的*周应订货物；已订 以多选方式显示并记录*周本次操作前未订货物'
        label = tk.Label(self.TabStrip1__Tab1,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=BOTTOM)
        self.TabStrip1.add(self.TabStrip1__Tab1, text='订货')
        
        # 标签页2 收货页
        self.TabStrip1__Tab2 = Frame(self.TabStrip1)
        l3 = tk.Label(self.TabStrip1__Tab2,text='已订',width=5,height=1)
        l3.place(x=60,y=10)
        self.klist1 = Listbox(self.TabStrip1__Tab2,width=20,height=45)
        self.klist1.place(x=10,y=30)
        l4 = tk.Label(self.TabStrip1__Tab2,text='未处理',width=5,height=1)
        l4.place(x=300,y=10)
        self.klist2 = Listbox(self.TabStrip1__Tab2,width=20,height=45,selectmode = MULTIPLE)
        self.klist2.place(x = 250,y = 30)
        l5 = tk.Label(self.TabStrip1__Tab2,text='本次已处理',width=8,height=1)
        l5.place(x=540,y=10)
        self.klist3 = Listbox(self.TabStrip1__Tab2,width=20,height=45,selectmode = MULTIPLE)
        self.klist3.place(x = 500,y = 30)
        b3 = tk.Button(self.TabStrip1__Tab2, text='收货', font=('Arial', 12), width=5, height=1, command=self.k_show_select)
        b3.place(x=420,y=280)
        b4 = tk.Button(self.TabStrip1__Tab2, text='取消', font=('Arial', 12), width=5, height=1, command=self.k_delect_select)
        b4.place(x=420,y=320)
        info = '应收 显示由订货的数据计算出的*周应收货物；已收 以多选方式显示并记录*周本次操作前未收货物'
        label = tk.Label(self.TabStrip1__Tab2,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=BOTTOM)
        # 点击 查看 按钮，显示 应订、未订
        b1 = tk.Button(self.TabStrip1__Tab2, text='查看', font=('Arial', 12), width=10, height=1, command=self.kontrol_checked)
        b1.place(x=750,y=25)
        # 点击 确认 按钮，将 已订 部分存入xls
        b2 = tk.Button(self.TabStrip1__Tab2, text='确认', font=('Arial', 12), width=10, height=1, command=self.kontrol_confirm)
        b2.place(x=750,y=65)
        b3 = tk.Button(self.TabStrip1__Tab2, text='投诉', font=('Arial', 12), width=10, height=1, command=self.kontrol_beschwe)
        b3.place(x=750,y=105)
        b4 = tk.Button(self.TabStrip1__Tab2, text='拍照', font=('Arial', 12), width=10, height=1, command=self.kontrol_foto)
        b4.place(x=750,y=145)
        
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.TabStrip1__Tab2,text='打开文件时间',width=12,font=('Arial',13))
        tl1.place(x=750,y=300)
        self.k_open_time = tk.StringVar()    
        tl2 = tk.Label(self.TabStrip1__Tab2, textvariable=self.k_open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=750,y=350)
        tl3 = tk.Label(self.TabStrip1__Tab2,text='当前文件时间',width=12,font=('Arial',13))
        tl3.place(x=750,y=400)
        self.k_file_time = tk.StringVar()    
        tl4 = tk.Label(self.TabStrip1__Tab2, textvariable=self.k_file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=750,y=450)
        self.TabStrip1.add(self.TabStrip1__Tab2, text='收货')
        
        # 标签页3 发票页
        self.TabStrip1__Tab3 = Frame(self.TabStrip1)
        l3 = tk.Label(self.TabStrip1__Tab3,text='已订',width=5,height=1)
        l3.place(x=60,y=10)
        self.rlist1 = Listbox(self.TabStrip1__Tab3,width=20,height=45)
        self.rlist1.place(x=10,y=30)
        l4 = tk.Label(self.TabStrip1__Tab3,text='未收',width=5,height=1)
        l4.place(x=300,y=10)
        self.rlist2 = Listbox(self.TabStrip1__Tab3,width=20,height=45,selectmode = MULTIPLE)    # selectmode = EXTENDED
        self.rlist2.place(x = 250,y = 30)
        l5 = tk.Label(self.TabStrip1__Tab3,text='本次已收',width=8,height=1)
        l5.place(x=540,y=10)
        self.rlist3 = Listbox(self.TabStrip1__Tab3,width=20,height=45,selectmode = MULTIPLE)
        self.rlist3.place(x = 500,y = 30)
        b3 = tk.Button(self.TabStrip1__Tab3, text='收票', font=('Arial', 12), width=5, height=1, command=self.r_show_select)
        b3.place(x=420,y=280)
        b4 = tk.Button(self.TabStrip1__Tab3, text='取消', font=('Arial', 12), width=5, height=1, command=self.r_delect_select)
        b4.place(x=420,y=320)
        info = '已订 显示由订货的数据计算出的*周应收发票；未收 以多选方式显示并记录*周本次操作前未收到的发票'
        label = tk.Label(self.TabStrip1__Tab3,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=BOTTOM)
        # 点击 查看 按钮，显示 应订、未订
        b1 = tk.Button(self.TabStrip1__Tab3, text='查看', font=('Arial', 12), width=10, height=1, command=self.rechnung_checked)
        b1.place(x=750,y=25)
        # 点击 确认 按钮，将 已订 部分存入xls
        b2 = tk.Button(self.TabStrip1__Tab3, text='确认', font=('Arial', 12), width=10, height=1, command=self.rechnung_confirm)
        b2.place(x=750,y=65)
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.TabStrip1__Tab3,text='打开文件时间',width=12,font=('Arial',13))
        tl1.place(x=750,y=300)
        self.r_open_time = tk.StringVar()    
        tl2 = tk.Label(self.TabStrip1__Tab3, textvariable=self.r_open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=750,y=350)
        tl3 = tk.Label(self.TabStrip1__Tab3,text='当前文件时间',width=12,font=('Arial',13))
        tl3.place(x=750,y=400)
        self.r_file_time = tk.StringVar()    
        tl4 = tk.Label(self.TabStrip1__Tab3, textvariable=self.r_file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=750,y=450)
        self.TabStrip1.add(self.TabStrip1__Tab3, text='发票')
        
        # 标签页4 录入页
        self.TabStrip1__Tab4 = Frame(self.TabStrip1)
        l3 = tk.Label(self.TabStrip1__Tab4,text='已订',width=5,height=1)
        l3.place(x=60,y=10)
        self.elist1 = Listbox(self.TabStrip1__Tab4,width=20,height=45)
        self.elist1.place(x=10,y=30)
        l4 = tk.Label(self.TabStrip1__Tab4,text='未录',width=5,height=1)
        l4.place(x=300,y=10)
        self.elist2 = Listbox(self.TabStrip1__Tab4,width=20,height=45,selectmode = MULTIPLE)    # selectmode = EXTENDED
        self.elist2.place(x = 250,y = 30)
        l5 = tk.Label(self.TabStrip1__Tab4,text='本次已录',width=8,height=1)
        l5.place(x=540,y=10)
        self.elist3 = Listbox(self.TabStrip1__Tab4,width=20,height=45,selectmode = MULTIPLE)
        self.elist3.place(x = 500,y = 30)
        b3 = tk.Button(self.TabStrip1__Tab4, text='确认', font=('Arial', 12), width=5, height=1, command=self.e_show_select)
        b3.place(x=420,y=280)
        b4 = tk.Button(self.TabStrip1__Tab4, text='取消', font=('Arial', 12), width=5, height=1, command=self.e_delect_select)
        b4.place(x=420,y=320)
        info = '已订 显示由订货的数据计算出的*周应录入货物；未录 以多选方式显示并记录*周本次操作前未录入的货物'
        label = tk.Label(self.TabStrip1__Tab4,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=BOTTOM)
        # 点击 查看 按钮，显示 应订、未订
        b1 = tk.Button(self.TabStrip1__Tab4, text='查看', font=('Arial', 12), width=10, height=1, command=self.eingeben_checked)
        b1.place(x=750,y=25)
        # 点击 确认 按钮，将 已订 部分存入xls
        b2 = tk.Button(self.TabStrip1__Tab4, text='录入', font=('Arial', 12), width=10, height=1, command=self.eingeben_confirm)
        b2.place(x=750,y=65)
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.TabStrip1__Tab4,text='打开文件时间',width=12,font=('Arial',13))
        tl1.place(x=750,y=300)
        self.e_open_time = tk.StringVar()    
        tl2 = tk.Label(self.TabStrip1__Tab4, textvariable=self.e_open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=750,y=350)
        tl3 = tk.Label(self.TabStrip1__Tab4,text='当前文件时间',width=12,font=('Arial',13))
        tl3.place(x=750,y=400)
        self.e_file_time = tk.StringVar()    
        tl4 = tk.Label(self.TabStrip1__Tab4, textvariable=self.e_file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=750,y=450)
        self.TabStrip1.add(self.TabStrip1__Tab4, text='录入')
        
        # 标签页5 投诉页
        self.TabStrip1__Tab5 = Frame(self.TabStrip1)
        l3 = tk.Label(self.TabStrip1__Tab5,text='需投诉',width=5,height=1)
        l3.place(x=60,y=10)
        self.wlist1 = Listbox(self.TabStrip1__Tab5,width=20,height=45)
        self.wlist1.place(x=10,y=30)
        l4 = tk.Label(self.TabStrip1__Tab5,text='未投诉',width=5,height=1)
        l4.place(x=300,y=10)
        self.wlist2 = Listbox(self.TabStrip1__Tab5,width=20,height=45,selectmode = MULTIPLE)    # selectmode = EXTENDED
        self.wlist2.place(x = 250,y = 30)
        l5 = tk.Label(self.TabStrip1__Tab5,text='本次已录',width=8,height=1)
        l5.place(x=540,y=10)
        self.wlist3 = Listbox(self.TabStrip1__Tab5,width=20,height=45,selectmode = MULTIPLE)
        self.wlist3.place(x = 500,y = 30)
        b3 = tk.Button(self.TabStrip1__Tab5, text='确认', font=('Arial', 12), width=5, height=1, command=self.w_show_select)
        b3.place(x=420,y=280)
        b4 = tk.Button(self.TabStrip1__Tab5, text='取消', font=('Arial', 12), width=5, height=1, command=self.w_delect_select)
        b4.place(x=420,y=320)
        info = '已订 显示由订货的数据计算出的*周已预定的货物；未投诉 以多选方式显示并记录*周本次操作前未进行投诉的货物'
        label = tk.Label(self.TabStrip1__Tab5,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=BOTTOM)
        # 点击 查看 按钮，显示 应订、未订
        b1 = tk.Button(self.TabStrip1__Tab5, text='查看', font=('Arial', 12), width=10, height=1, command=self.beschwer_checked)
        b1.place(x=750,y=25)
        # 点击 确认 按钮，将 已订 部分存入xls
        b2 = tk.Button(self.TabStrip1__Tab5, text='确认投诉', font=('Arial', 12), width=10, height=1, command=self.beschwer_confirm)
        b2.place(x=750,y=65)
        b3 = tk.Button(self.TabStrip1__Tab5, text='无需投诉', font=('Arial', 12), width=10, height=1, command=self.beschwer_cancel)
        b3.place(x=750,y=105)
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.TabStrip1__Tab5,text='打开文件时间',width=12,font=('Arial',13))
        tl1.place(x=750,y=300)
        self.w_open_time = tk.StringVar()    
        tl2 = tk.Label(self.TabStrip1__Tab5, textvariable=self.w_open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=750,y=350)
        tl3 = tk.Label(self.TabStrip1__Tab5,text='当前文件时间',width=12,font=('Arial',13))
        tl3.place(x=750,y=400)
        self.w_file_time = tk.StringVar()    
        tl4 = tk.Label(self.TabStrip1__Tab5, textvariable=self.w_file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=750,y=450)
        self.TabStrip1.add(self.TabStrip1__Tab5, text='投诉')
        
        # 标签页6 查看页
        self.TabStrip1__Tab6 = Frame(self.TabStrip1)
        self.status = tk.StringVar()
        # lab = tk.Label(self.TabStrip1__Tab6,textvariable=self.status,font=('Arial',12),height=40,justify = LEFT)
        # lab.place(x=60,y=10)
        self.scr = scrolledtext.ScrolledText(self.TabStrip1__Tab6, width=85, height=35,font=("隶书",14))    # 加入滚动条以输出多行文本
        self.scr.place(x=30,y=60)
        but = tk.Button(self.TabStrip1__Tab6,text='查看',font=('Arial',12),width=10,height=1,command=self.checked)
        but.place(x=600,y=10)
        but1 = tk.Button(self.TabStrip1__Tab6,text='说明',font=('Arial',12),width=10,height=1,command=self.info)
        but1.place(x=700,y=10)
        but2 = tk.Button(self.TabStrip1__Tab6,text='更新总览',font=('Arial',12),width=10,height=1,command=self.updata)
        but2.place(x=800,y=10)
        self.TabStrip1.add(self.TabStrip1__Tab6, text='查看')
        
        self.TabStrip1.bind('<<NotebookTabChanged>>',self.tab_change)
        # TODO：
        # 标签页7 发订单o
        # 订货、发订单、收发票等加人名按钮，以区分操作者
        # 原件页->收到原件 判定：k
        
 
class Application(Application_ui):
    #这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    # TODO: 多标签页代码汇总，以节省运算空间
    '''    
        if flag == 1:
            woche = int(self.week.get())
            if woche < 10:
                woche = '0' + str(woche)
            else: woche = str(woche)
        else: 
            woche = int(self.week.get()) - 1
            if woche < 10:
                woche = '0' + str(woche)
            else: woche = str(woche)
        local_time = time.strftime("%H:%M:%S", time.localtime())
        self.open_time.set(local_time)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if flag == 1:
            self.var.set('预定')
        elif flag == 2:
            self.var.set('收货')
        elif flag == 3:
            self.var.set('发票')
        elif flag == 4:
            self.var.set('录入')
        elif flag == 5:
            self.var.set('投诉')
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.askquestion(title='Error',message='请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            if 'b' not in self.dframe.at[self.cbxf.get(),self.cbxl.get()]:
                message = 'KW' + woche + self.cbxf.get() + self.cbxl.get() + '未订货'
                tk.messagebox.askquestion(title='Warn',message=message)
    '''
    def tab_change(self,*args):
        # 切换标签页清空列表
        self.blist1.delete(0,END)
        self.blist2.delete(0,END)
        self.blist3.delete(0,END)
        self.klist1.delete(0,END)
        self.klist2.delete(0,END)
        self.klist3.delete(0,END)
        self.rlist1.delete(0,END)
        self.rlist2.delete(0,END)
        self.rlist3.delete(0,END)
        self.elist1.delete(0,END)
        self.elist2.delete(0,END)
        self.elist3.delete(0,END)
        self.wlist1.delete(0,END)
        self.wlist2.delete(0,END)
        self.wlist3.delete(0,END)
        
    def __init__(self, master=None):
        Application_ui.__init__(self, master)
    def get_FileModifyTime(self,filePath):
        # 获取文件更改时间
        t = os.path.getmtime(filePath)
        return time.localtime(t)

    def checked(self):
        self.var.set('查看')
        self.scr.delete(1.0, END)
        s = ''
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        s = '查看时间：' + local_time + '\n'
        woche = int(self.week.get())
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            # 输出总览表
            # TODO: 保证格式 打印 s = s + str(self.dframe)
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            info = str(self.dframe.at[i,self.cbxl.get()]) 
            s = s + 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '  详细信息:\n\n'
            if 'b' in info:
                s = s + '已订货 '
                if 'r' in info:
                    s = s + '已有发票 '
                else: s = s + '未收发票 '
                if 'k' in info:
                    s = s + '已收货 '
                    if 'z' in info: 
                        s = s + '有拍照确认 '
                    else: s = s + '无拍照确认 '
                    if 'y' in info: 
                        s = s + '点货后需投诉 '
                        if 'w' in info:
                            s = s + '已投诉 '
                        elif 'i' in info:
                            s = s + '不需要投诉'
                    else: s = s + '一切正常'
                else: s = s + '未收货'
            else: s = s + '未订货'
        elif self.cbxf.get():
            s = s + 'KW' + woche + ' ' + self.cbxf.get() + '  详细信息:\n\n'
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                if col == 'ID':continue
                elif col == 'Unnamed: 0': continue
                # s = s + col + (10-len(col)) * 2 * ' ' +': '
                c = '%10s' % col
                s = s + c + ' :'
                info = str(self.dframe.at[i,col])
                if 'b' in info:
                    s = s + '已订货 '
                    if 'r' in info:
                        s = s + '已有发票 '
                    else: s = s + '未收发票'
                    if 'k' in info:
                        s = s + '已收货 '
                        if 'z' in info: 
                            s = s + '有拍照确认 '
                        else: s = s + '无拍照确认 '
                        if 'y' in info: 
                            s = s + '点货后需投诉 '
                            if 'w' in info:
                                s = s + '已投诉 '
                            elif 'i' in info:
                                s = s + '不需要投诉 '
                        else: s = s + '一切正常 '
                    else: s = s + '未收货 '
                    
                else: s = s + '未订货 '
                s = s + '\n'
        else:
            s = s + 'KW' + woche + ' ' + self.cbxl.get() + '  详细信息:\n\n'
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                s = s + self.dframe.at[i,'ID'] + (10-len(self.dframe.at[i,'ID'])) * ' ' + ': '
                if 'b' in info:
                    s = s + '已订货 '
                    if 'r' in info:
                        s = s + '已有发票 '
                    else: s = s + '未收发票 '
                    if 'k' in info:
                        s = s + '已收货 '
                        if 'y' in info: 
                            s = s + '点货后需投诉 '
                        if 'w' in info:
                            s = s + '已投诉 '
                        elif 'i' in info:
                            s = s + '不需要投诉'
                        else: s = s + '一切正常'
                    else: s = s + '未收货'
                else: s = s + '未订货 '
                s = s + '\n'
        self.scr.insert(END,s)
        # self.status.set(s)
    def info(self):
        self.var.set('info')
        self.scr.delete(1.0, END)
        files = 'INFO.txt'
        s = '现有功能及使用说明\n'
        with open(files,'r',encoding='utf-8') as f1:
            line = f1.readline()
            while line:
                s = s + line
                line = f1.readline()
        # self.status.set(s)
        self.scr.insert(END,s)
    def updata(self):
        week = datetime.now().isocalendar()[1]
        if week < 10:
            week = '0' + str(week)
        else: week = str(week)
        file = 'KW' + week + '.xlsx'
        
        today = datetime.now().weekday() # 周一为0
        now = int(time.time())
        FileTime = os.path.getmtime(file) # 时间戳
        file_time = time.mktime(time.strftime("%Y-%m-%d %H:%M",FileTime))
        if today > 3 and (now - FileTime > 5):
            TableReader().Updata_to_LF(week) # 更新总览表
        file = 'KW' + week + '.xlsx'
        
        # path = 'H:\\py\\test\\goasia'
        # file_list = os.listdir(path)
        file_list = os.listdir('.')
        if file not in file_list: 
            TableReader().Writer()

    def GetINFO(self,file):
        self.dframe = TableReader().Reader(file) # 分店索引从0开始
    
    def bestell_checked(self):
        # 点击 查看 后，从供货商文档中获取列表，对比总览表中的数据，显示本周应订
        # 显示n周应到货、未到货信息
        self.var.set('预定')
        self.blist1.delete(0,END)
        self.blist2.delete(0,END)
        self.blist3.delete(0,END)
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.b_open_time.set(local_time)
        woche = int(self.week.get())
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.b_file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            if 's' not in str(self.dframe.at[i,self.cbxl.get()]):
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '无须订货'
                tk.messagebox.showwarning(title='Warn',message=message)
            else:
                self.blist2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                if col == 'ID': continue
                elif col == 'Unnamed: 0': continue
                info = str(self.dframe.at[i,col])
                if 's' in info:
                    self.blist1.insert(END,col)
                if 'b' not in info:
                # if 'b' not in info and 's' in info:
                    self.blist2.insert(END,col)
        else:
            # 确认供货商，以门店查询
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                if 's' in info: # 判断是否应订
                    # print(self.dframe.at[i,'ID'])
                    self.blist1.insert(END,self.dframe.at[i,'ID'])
                if 'b' not in info:
                    self.blist2.insert(END,self.dframe.at[i,'ID'])
    def bestell_confirm(self):
        a = self.blist3.size()
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='预定时间非当前周，请确认是否仍要继续')
        print(a)
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.blist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.blist3.get(i)]) for i in range(self.blist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.blist3.get(i)]) for i in range(self.blist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'b')
    def b_show_select(self,*args):
        a = self.blist2.size()
        for i in range(a):
            if(self.blist2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                # 删除未订表中的相应项，并将其加入已订表中
                self.blist3.insert(END,self.blist2.get(a-1-i))
                self.blist2.delete(a-1-i)
    def b_delect_select(self,*args):
        a = self.blist3.size()
        for i in range(a):
            if(self.blist3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                # 删除已订表中的相应项，并将其加入未订表中
                self.blist2.insert(END,self.blist3.get(a-1-i))
                self.blist3.delete(a-1-i)
            
    def kontrol_checked(self):
        self.var.set('收货')
        self.klist1.delete(0,END)
        self.klist2.delete(0,END)
        self.klist3.delete(0,END)
        woche = int(self.week.get())
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.k_open_time.set(local_time)
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.k_file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
            # tk.messagebox.askquestion(title='Error',message='请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            if 'b' not in str(self.dframe.at[i,self.cbxl.get()]):
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '未订货'
                tk.messagebox.askquestion(title='Warn',message=message)
            else:
                self.klist2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                info = str(self.dframe.at[i,col])
                if 'b' in info:
                    self.klist1.insert(END,col)
                    if 'k' not in info:
                        self.klist2.insert(END,col)
        else:
            # 确认供货商，以门店查询
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                if 'b' in info: # 判断是否应订
                    self.klist1.insert(END,self.dframe.at[i,'ID'])
                    if 'k' not in info:
                        self.klist2.insert(END,self.dframe.at[i,'ID'])
    def kontrol_confirm(self):
        a = self.klist3.size()
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='收货时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.klist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'k')
    def k_show_select(self,*args):
        a = self.klist2.size()
        for i in range(a):
            if(self.klist2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                # 删除未订表中的相应项，并将其加入已订表中
                self.klist3.insert(END,self.klist2.get(a-1-i))
                self.klist2.delete(a-1-i)
    def k_delect_select(self,*args):
        a = self.klist3.size()
        for i in range(a):
            if(self.klist3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.klist2.insert(END,self.klist3.get(a-1-i))
                self.klist3.delete(a-1-i)
    def kontrol_beschwe(self):
        a = self.klist3.size()
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='收货时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.klist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'ky') # 标记点货、投诉
    def kontrol_foto(self):
        a = self.klist3.size()
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='收货时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.klist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.klist3.get(i)]) for i in range(self.klist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'z') # 标记拍照登记

    def r_show_select(self,*args):
        a = self.rlist2.size()
        for i in range(a):
            if(self.rlist2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.rlist3.insert(END,self.rlist2.get(a-1-i))
                self.rlist2.delete(a-1-i)
    def r_delect_select(self,*args):
        a = self.rlist3.size()
        for i in range(a):
            if(self.rlist3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.rlist2.insert(END,self.rlist3.get(a-1-i))
                self.rlist3.delete(a-1-i)
    def rechnung_checked(self):
        self.var.set('发票')
        self.rlist1.delete(0,END)
        self.rlist2.delete(0,END)
        self.rlist3.delete(0,END)
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.r_open_time.set(local_time)
        woche = int(self.week.get())
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.r_file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            if 'b' not in str(self.dframe.at[i,self.cbxl.get()]):
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '未订货'
                tk.messagebox.showwarning('Error',message)
            else:
                self.rlist2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                info = str(self.dframe.at[i,col])
                if 'b' in info:
                    self.rlist1.insert(END,col)
                    if 'r' not in info:
                        self.rlist2.insert(END,col)
        else:
            # 确认供货商，以门店查询
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                if 'b' in info: # 判断是否应订
                    self.rlist1.insert(END,self.dframe.at[i,'ID'])
                    if 'r' not in info:
                        self.rlist2.insert(END,self.dframe.at[i,'ID'])
    def rechnung_confirm(self):
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='收货时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.rlist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.rlist3.get(i)]) for i in range(self.rlist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.rlist3.get(i)]) for i in range(self.rlist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'r')

    def e_show_select(self,*args):
        a = self.elist2.size()
        for i in range(a):
            if(self.elist2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.elist3.insert(END,self.elist2.get(a-1-i))
                self.elist2.delete(a-1-i)
    def e_delect_select(self,*args):
        a = self.elist3.size()
        for i in range(a):
            if(self.elist3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.elist2.insert(END,self.elist3.get(a-1-i))
                self.elist3.delete(a-1-i)
    def eingeben_checked(self):
        self.var.set('录入')
        self.elist1.delete(0,END)
        self.elist2.delete(0,END)
        self.elist3.delete(0,END)
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.e_open_time.set(local_time)
        woche = int(self.week.get())
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.e_file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            if 'b' not in str(self.dframe.at[i,self.cbxl.get()]):
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '未订货'
                tk.messagebox.showwarning('Error',message)
            else:
                self.elist2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                info = str(self.dframe.at[i,col])
                if 'b' in info:
                    self.elist1.insert(END,col)
                    if 'e' not in info:
                        self.elist2.insert(END,col)
        else:
            # 确认供货商，以门店查询
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                if 'b' in info: # 判断是否应订
                    self.elist1.insert(END,self.dframe.at[i,'ID'])
                    if 'e' not in info:
                        self.elist2.insert(END,self.dframe.at[i,'ID'])
    def eingeben_confirm(self):
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='操作时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.elist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.elist3.get(i)]) for i in range(self.elist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.elist3.get(i)]) for i in range(self.elist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'e')

    def w_show_select(self,*args):
        a = self.wlist2.size()
        for i in range(a):
            if(self.wlist2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.wlist3.insert(END,self.wlist2.get(a-1-i))
                self.wlist2.delete(a-1-i)
    def w_delect_select(self,*args):
        a = self.wlist3.size()
        for i in range(a):
            if(self.wlist3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                self.wlist2.insert(END,self.wlist3.get(a-1-i))
                self.wlist3.delete(a-1-i)
    def beschwer_checked(self):
        self.var.set('投诉')
        self.wlist1.delete(0,END)
        self.wlist2.delete(0,END)
        self.wlist3.delete(0,END)
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.w_open_time.set(local_time)
        woche = int(self.week.get())
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FileTime = self.get_FileModifyTime(file)
        self.w_file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        self.GetINFO(file)
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            i = self.FLd[self.cbxf.get()]
            if 'y' not in str(self.dframe.at[i,self.cbxl.get()]):
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '不需要投诉'
                tk.messagebox.showwarning('Error',message)
            else:
                self.wlist2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            i = self.FLd[self.cbxf.get()]
            for col in self.dframe.columns.values.tolist():
                info = str(self.dframe.at[i,col])
                if 'y' in info:
                    self.wlist1.insert(END,col)
                    if 'w' not in info and 'i' not in info:
                        self.wlist2.insert(END,col)
        else:
            # 确认供货商，以门店查询
            for i in range(self.dframe.shape[0]):
                info = str(self.dframe[self.cbxl.get()][i])
                if 'y' in info: # 判断是否应投诉
                    self.wlist1.insert(END,self.dframe.at[i,'ID'])
                    if 'w' not in info and 'i' not in info:
                        self.wlist2.insert(END,self.dframe.at[i,'ID'])
    def beschwer_confirm(self):
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='操作时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.wlist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.wlist3.get(i)]) for i in range(self.wlist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.wlist3.get(i)]) for i in range(self.wlist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'w')
    def beschwer_cancel(self)  :
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='操作时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        if woche < 10:
            woche = '0' + str(woche)
        else: woche = str(woche)
        file = 'KW' + woche + '.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.wlist3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FLd[self.cbxf.get()])
            {LF_updata.append(self.LFd[self.wlist3.get(i)]) for i in range(self.wlist3.size())}
        else:
            LF_updata.append(self.LFd[self.cbxl.get()])
            {FL_updata.append(self.FLd[self.wlist3.get(i)]) for i in range(self.wlist3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,'i')  # 标记i（in ordnung）

if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()
    
    # TODO:
    # 应订update 建新UI处理
    # UI：保留应订；提供自定义选项，允许进行加减操作，生成订货表
    # 建立新表：按分店、供货商查询订货周期
    # 查看上一周的订货信息
    # 创建订货表： 订货时间提醒、某供货商送货的固定路线->门店(现：供货商、时间，表格内容门店)
    
    # 生成到货表（需要建立新表记录各供货商、门店到货时间信息）
    # 根据订货表以供货商为单位建立到货表（供货商、时间）->以店为单位发送到各店
    # 到货表： 根据分店、供货商到货时间、板数
    # 当周到（算法实现？x）、下周到(W标记固定周几)、几天后到(T标记订货后几天到)；标记订货时间？how计算T状况的到货情况
