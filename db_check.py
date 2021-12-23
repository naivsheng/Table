'''
# -*- coding: UTF-8 -*-
# __Author__: Yingyu Wang
# __date__: 23.12.2021
# __Version__: 2.01 更新FL、LF表，总览表中获取准确信息
'''
import tkinter as tk
from tkinter.font import Font
from tkinter.ttk import *
from tkinter.messagebox import *
import os
import pandas as pd
from TableReader import TableReader
import time
from datetime import datetime
from tkinter import scrolledtext, END, Toplevel
import sys
from threading import Timer

class Application_ui(Frame):
    # 这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    # 允许向表格填充自定义内容
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('订货录入辅助')
        self.master.geometry('1200x800')
        filePath = os.getcwd()
        self.FL = ['']
        df = pd.read_excel('Filialen.xlsx',header = 0)
        self.FL.extend(df.iloc[:,0])
        self.LF = ['']
        df = pd.read_excel('Lieferant.xlsx',header = None)
        self.LF.extend(df.iloc[0,1:])
        self.Lieferant = pd.read_excel('Lieferant.xlsx', header=0, index_col=0)   # 获取订货周期、最后一次订货时间
        self.choose_status = ['_订货','_发票','_账单','_到货','_入货','_传真','_投诉','_原件']
        self.createWidgets()
    
    def createWidgets(self):
        '''
            # 创建主窗口：以单选框形式进行操作转换
            状态栏公用
            文件打开、更改时间记录/刷新提醒
            TODO：gmailAPI读取：确认发票信息（从pdf附件读取供货商、收货门店信息）
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
        # 建立下拉选框，选择分店
        l2 = tk.Label(self.top,text='分店',width=5,height=1)
        l2.place(x=240,y=10)
        valueF = tk.StringVar()
        self.cbxf = tk.ttk.Combobox(self.top, width = 10, height = 20, textvariable = valueF,state='readonly') #, postcommand = self.show_select)
        self.cbxf['value'] = self.FL
        self.cbxf.place(x=280,y=10)
        l6 = tk.Label(self.top,text='供应商',width=8,height=1)
        l6.place(x=400,y=10)
        valueL = tk.StringVar()
        self.cbxl = tk.ttk.Combobox(self.top,width = 25, height = 20, textvariable = valueL,state='readonly')
        self.cbxl["value"] = self.LF
        self.cbxl.place(x=480,y=10)
        self.choose = tk.IntVar()
        for i in range(len(self.choose_status)):
            tk.Radiobutton(self.top,text=self.choose_status[i],variable=self.choose,value=i,command=self.Operationen).place(x=50,y=i*50+100)
        l_mark = tk.Label(self.top,text='填充标记',width=10,height=1,font=('Arial', 12))
        l_mark.place(x=1000,y=50)
        self.mark = tk.StringVar()
        entry_mark = tk.Entry(self.top,textvariable=self.mark,width=10,font=('Arial', 12))
        entry_mark.place(x=1000,y=80)
        b1 = tk.Button(self.top, text='查看', font=('Arial', 12), width=10, height=1, command=self.bestell_checked)
        b1.place(x=1000,y=150)
        b2 = tk.Button(self.top, text='确认', font=('Arial', 12), width=10, height=1, command=self.bestell_confirm)
        b2.place(x=1000,y=200)
        # 记录文件打开时间，用以判断是否已有更改，提醒操作人刷新最新数据
        tl1 = tk.Label(self.top,text='打开当前文件的时间',width=20,font=('Arial',13))
        tl1.place(x=950,y=400)
        self.open_time = tk.StringVar()    
        tl2 = tk.Label(self.top, textvariable=self.open_time,font=('Arial', 14), width=12, height=1)
        tl2.place(x=970,y=450)
        tl3 = tk.Label(self.top,text='当前文件的更改时间',width=20,font=('Arial',13))
        tl3.place(x=950,y=500)
        self.file_time = tk.StringVar()    
        tl4 = tk.Label(self.top, textvariable=self.file_time,font=('Arial', 14), width=12, height=1)
        tl4.place(x=970,y=550)
        tl5 = tk.Label(self.top,text='当前系统时间',width=20,font=('Arial',13))
        tl5.place(x=950,y=300)
        self.actuelle_time = tk.StringVar()    
        tl6 = tk.Label(self.top, textvariable=self.actuelle_time,font=('Arial', 14), width=12, height=1)
        tl6.place(x=970,y=350)
        l3 = tk.Label(self.top,text='应标记',width=5,height=1)
        l3.place(x=240,y=50)
        sb = Scrollbar(self.top) # 给列表增加滚动条，以防过多数据
        self.list1 = tk.Listbox(self.top,width=25,height=45,yscrollcommand=sb.set)
        self.list1.place(x=160,y=70)
        sb.config(command=self.list1.yview)
        l4 = tk.Label(self.top,text='未标记',width=5,height=1)
        l4.place(x=520,y=50)
        self.list2 = tk.Listbox(self.top,width=25,height=45,selectmode = tk.MULTIPLE)
        self.list2.place(x = 450,y = 70)
        l5 = tk.Label(self.top,text='本次已操作',width=8,height=1)
        l5.place(x=820,y=50)
        self.list3 = tk.Listbox(self.top,width=25,height=45,selectmode = tk.MULTIPLE)
        self.list3.place(x = 750,y = 70)
        b3 = tk.Button(self.top, text='—>', font=('Arial', 12), width=5, height=1, command=self.show_select)
        b3.place(x=670,y=300)
        b4 = tk.Button(self.top, text='<—', font=('Arial', 12), width=5, height=1, command=self.delect_select)
        b4.place(x=670,y=350)
        info = '应标记 显示由之前的数据计算出的*周应进行的操作；未标记 以多选方式显示并记录*周本次操作前未进行的操作'
        label = tk.Label(self.top,text = info, fg='green',font=('Arial',12),width=500)
        label.pack(side=tk.BOTTOM)
        self.top.iconbitmap('goasia.ico')
        b_info = tk.Button(self.top, text='说明', font=('Arial', 12), width=10, height=1, command=self.Info)
        b_info.place(x=800,y=10)
        self.status = False
        self.refresh_data()
         
class Application(Application_ui):
    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    # 多标签页代码汇总，以节省运算空间
    def __init__(self, master=None):
        Application_ui.__init__(self, master)
    def tab_change(self,*args):
        '''切换标签页清空列表'''
        self.list1.delete(0,END)
        self.list2.delete(0,END)
        self.list3.delete(0,END)

    def refresh_data(self,filePath=None):
        '''定时刷新文件的更改时间'''
        self.actuelle_time.set(time.strftime("%m-%d %H:%M:%S",time.localtime(int(time.time()))))
        self.timer = self.after(1800000,self.refresh_data) # 每10分钟刷新600000
        if self.status:
            tk.messagebox.showinfo('提示','请点击查看以刷新数据')
        else:
            self.status = True
            tk.messagebox.showinfo('提示','本程序将进行每十分钟弹窗提醒，不会对操作内容产生影响')
        # t = os.path.getmtime(filePath)
        # self.file_time.set(time.strftime("%m-%d %H:%M:%S",time.localtime(t)))
    
    def Operationen(self):
        '''通过单选项更改订货表单元格位置'''
        self.tab_change()
        return True
    
    def get_FileModifyTime(self,filePath):
        '''获取文件更改时间'''
        t = os.path.getmtime(filePath)
        self.file_time.set(time.strftime("%m-%d %H:%M:%S",time.localtime(t)))
        # return time.localtime(t)

    def checked(self):
        '''查看总览信息'''
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
    def Info(self):
        '''弹窗显示说明'''
        self.var.set('info')
        top = Toplevel()
        top.geometry('795x690')
        top.title('Info')
        self.scr = scrolledtext.ScrolledText(top, width=76, height=35,font=("隶书",14),bg='whitesmoke')    # 加入滚动条以输出多行文本
        self.scr.place(x=10,y=10)
        # self.scr.delete(1.0, END)
        files = 'INFO.txt'
        s = '现有功能及使用说明\n'
        with open(files,'r',encoding='utf-8') as f1:
            line = f1.readline()
            while line:
                s = s + line
                line = f1.readline()
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
        # 点击 查看 后，从供货商文档中获取列表，对比总览表中的数据，显示操作信息
        self.var.set(self.choose_status[self.choose.get()])
        self.tab_change()
        self.status = True
        local_time = time.strftime("%m-%d %H:%M:%S", time.localtime())
        self.open_time.set(local_time)
        nextwoche = str(int(self.week.get()) + 1).zfill(2)
        woche = self.week.get().zfill(2)
        file = f'KW{woche} Bestellung KW{nextwoche} Lieferung Übersicht.xlsx'
        self.GetINFO(file)
        # 更改数据，创建可查询的dataframe
        # 获取门店列表、供应商列表
        new_col = []
        LF_new = ['']
        for col in list(self.dframe): # 遍历列名
            a = str(self.dframe[col][1]) if str(self.dframe[col][1]) != 'nan' else a
            if str(self.dframe[col][1]) != 'nan': LF_new.append(a) 
            new_col.append(f'{a}_{self.dframe[col][2]}')
        self.dframe.columns = new_col
        if LF_new != self.LF:
            self.LF = LF_new
            TableReader().Update_new_LF(LF_new)
        checkdf = self.dframe[3:]
        FL_new = list(checkdf.index)
        for i in list(checkdf.index):
            if not str(i).split(' ')[0].isdigit():
                FL_new.remove(i)
        FL_new.insert(0,'')
        if FL_new != self.FL:
            self.FL = FL_new
            TableReader().Update_new_FL(FL_new)
        # FileTime = self.get_FileModifyTime(file)
        # self.file_time.set(time.strftime("%m-%d %H:%M:%S",FileTime))
        # print(self.Lieferant)   # 上次订货时间
        # 获取并显示信息
        if not self.cbxf.get() and not self.cbxl.get():
            self.var.set('Error')
            tk.messagebox.showwarning('Error','请选择分店或供应商')
        elif self.cbxf.get() and self.cbxl.get():
            fl = self.FL.index(self.cbxf.get())  
            lf = self.LF.index(self.cbxl.get())
            # TODO 判断是否应订
            # lastbestellung = self.Lieferant.loc[self.cbxf.get(),self.cbxl.get()]
            # zeitraum = self.Lieferant.loc[self.cbxf.get(),'订货周期']
            # struct_day = time.strptime(lastbestellung, "%Y-%m-%d")
            # week = datetime.date(struct_day[0],struct_day[1],struct_day[2]).isocalendar()[1]    # Y,m,d 获取订货日期所在的KW
            # if week > int(self.week.get()):
            #     pass
            # if monat == 1 and week > 10:
            #     s = get + '在' + str(year - 1) + '的第' + str(week) + '周' 
            
            row = self.cbxf.get()
            col = self.cbxl.get() + self.choose_status[self.choose.get()]
            if str(self.dframe.at[row,col]) != 'nan':
                message = 'KW' + woche + ' ' + self.cbxf.get() + '  ' + self.cbxl.get() + '已标记'
                tk.messagebox.showwarning(title='Warn',message=message)
            else:
                self.list2.insert(END,self.cbxl.get())
        elif self.cbxf.get():
            # 确认门店(行)，以供货商(列)形式查询
            # TODO 订货：应订计算； 发票等：订货判定
            row = self.cbxf.get()
            del LF_new[0]
            LF_new = [i + self.choose_status[self.choose.get()] for i in LF_new]
            if self.choose.get() == 0:
                # TODO 计算应订
                check = ''
            elif self.choose.get() < 4:
                check = '_订货'
            else:
                check = '_到货'
            for col in LF_new:
                info = str(self.dframe.at[row,col])
                if not check:
                    # list1 暂时置空
                    pass
                else:
                    check_info = str(self.dframe.at[row,col.split('_')[0] + check])
                    if check_info != 'nan': # 应进行操作
                        self.list1.insert(END,col.split('_')[0])
                if info == 'nan':
                    self.list2.insert(END,col.split('_')[0])
        else:
            # 确认供货商，以门店查询
            col = self.cbxl.get() + self.choose_status[self.choose.get()]
            if self.choose.get() == 0:
                # list1 置空，list2填充
                check = ''
            elif self.choose.get() < 4:
                check = self.cbxl.get() + '_订货'
            else:
                check = self.cbxl.get() + '_到货'
            for i in self.FL:
                if i == '':continue
                info = str(self.dframe[self.cbxl.get() + self.choose_status[self.choose.get()]][i])
                if not check:
                    pass
                else:
                    if str(self.dframe[check][i]) != 'nan':
                        self.list1.insert(END,i)
                if info == 'nan':
                    self.list2.insert(END,i)
    def bestell_confirm(self):
        '''把选定项、标记写入表格'''
        if not self.mark.get():
            tk.messagebox.showinfo('提示','请输入操作标记')
            return True
        a = self.list3.size()
        woche = int(self.week.get())
        now = time.strftime("%W")
        if woche < int(now):
            a = tk.messagebox.askquestion(title='Warning',message='预定时间非当前周，请确认是否仍要继续')
        if not a:
            # 用户点击取消，本次操作不保存
            return False
        nextwoche = str(int(self.week.get()) + 1).zfill(2)
        woche = self.week.get().zfill(2)
        file =  f'KW{woche} Bestellung KW{nextwoche} Lieferung Übersicht.xlsx'
        FL_updata = []
        LF_updata = []
        if not self.list3:
            # 没有选择，直接返回
            return True
        if self.cbxf.get():
            FL_updata.append(self.FL.index(self.cbxf.get())) # 返回行号
            {LF_updata.append(list(self.dframe.columns).index(self.list3.get(i) + self.choose_status[self.choose.get()])) for i in range(self.list3.size())}
        else:
            LF_updata.append(list(self.dframe.columns).index(self.cbxl.get() + self.choose_status[self.choose.get()]))
            {FL_updata.append(self.FL.index(self.list3.get(i))) for i in range(self.list3.size())}
        TableReader().Updata(file,FL_updata,LF_updata,self.mark.get())
        # TODO 清除信息？ 停止计时？
        # self.status = False
        # self.after_cancel(self.timer) # 停止计时
        
    def show_select(self,*args):
        a = self.list2.size()
        for i in range(a):
            if(self.list2.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                # 删除未订表中的相应项，并将其加入已订表中
                self.list3.insert(END,self.list2.get(a-1-i))
                self.list2.delete(a-1-i)
    
    def delect_select(self,*args):
        a = self.list3.size()
        for i in range(a):
            if(self.list3.select_includes(a-1-i)) == True: # 判断是否选中list中的数据
                # 删除已订表中的相应项，并将其加入未订表中
                self.list2.insert(END,self.list3.get(a-1-i))
                self.list3.delete(a-1-i)

if __name__ == "__main__":
    # df = pd.read_excel('KW29 Bestellung KW30 Lieferung Übersicht.xlsx', header=0)
    # print(df.iat[3,3]) # 读取D5的数据
    top = tk.Tk()
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
