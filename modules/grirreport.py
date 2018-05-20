# -*- coding: UTF-8 -*-
import pandas as pd
from tkinter import *
from tkinter import filedialog
from datetime import datetime
from datetime import timedelta
import os
import winreg 
import xlrd
import calendar
class GRIRReport():
    def __init__(self):
        self.root=Tk()
        self.root.title("GRIRReport")
        self.w1=Frame(height=200,width=500)
        self.w2=Frame(height=30,width=500)
        self.w4=Frame(self.w2,height=30,width=65)
        self.w5=Frame(self.w2,height=30,width=370)
        self.w6=Frame(self.w2,height=30,width=65)
        self.w1.grid_propagate(0)
        self.w1.grid(row=0,column=0,padx=2, pady=5)
        self.w2.grid(row=2)
        self.w4.pack(side='left')
        self.w5.pack(side='left')
        self.w6.pack(side='right')

        self.scrollbar = Scrollbar(self.w1)
        self.scrollbar.pack( side = RIGHT, fill=Y )
        self.context_text=Text(self.w1, yscrollcommand = self.scrollbar.set)
        self.context_text_index = 1.0
        self.context_text.pack(side=LEFT, fill=Y)
        self.openfile_btn=Button(self.w4, text='打开文件', command=self.openfile)
        self.openfile_btn.pack(side='left')
        self.close_btn=Button(self.w6, text='关闭软件', command=self.close)
        self.close_btn.pack(side='right')
        self.root.mainloop()

    def processexcel(self,in_file_name):
        sheet_name = "Sheet1"
        temp = in_file_name.split('/')
        out_file_name = self.get_desktop() + '\DATA\GRIRReport-output-' + temp[len(temp)-1]
        in_col_filters = ['Order', 'Supplier', 'Supplier Name', 'Entity', 'Account',
                          '0-30 Days CNY', '31-60 Days CNY', '61-90 Days CNY', '91-180 Days CNY',
                          '181-360 Days CNY', 'Over 360 Days CNY']
        col_methods = {'Supplier': 'first', 'Supplier Name': 'first',
                       'Entity': 'first', 'Account': 'first', '0-30 Days CNY': 'sum', '31-60 Days CNY': 'sum',
                       '61-90 Days CNY': 'sum', '91-180 Days CNY': 'sum', '181-360 Days CNY': 'sum', 'Over 360 Days CNY': 'sum'}
        row_sum_names = ['0-30 Days CNY', '31-60 Days CNY', '61-90 Days CNY', '91-180 Days CNY',
                         '181-360 Days CNY', 'Over 360 Days CNY']
        col_sum_names = ['0-30 Days CNY', '31-60 Days CNY', '61-90 Days CNY', '91-180 Days CNY',
                         '181-360 Days CNY', 'Over 360 Days CNY', 'Amount']
        out_col_filters = ['Supplier', 'Supplier Name', 'Entity', 'Account',
                           '0-30 Days CNY', '31-60 Days CNY', '61-90 Days CNY', '91-180 Days CNY',
                           '181-360 Days CNY', 'Over 360 Days CNY', 'Amount']
        out_col_names = {'Supplier':'Vendor Code', 'Supplier Name':'Supplier'}
        xls_file = pd.ExcelFile(in_file_name)
        try:
            raw = xls_file.parse(sheet_name, fill_value=0)
        except Exception as e:
            self.context_text.insert(self.context_text_index, '文件选择错误或sheet重命名为Sheet1！！！\n')
            self.context_text_index += 1
            return '请重新打开文件！！！'
        try:
            table = raw[in_col_filters]
        except Exception as e:
            self.context_text.insert(self.context_text_index, '文件头数据不匹配，请检查！！！\n')
            self.context_text_index += 1
            return '请重新打开文件！！！'
        table.rename(columns = {'Order':'Po Number'}, inplace=True)
        writer = pd.ExcelWriter(out_file_name)
        for (Entity, Account), group in table.groupby(['Entity', 'Account']):
            name = str(Entity) + " " + str(Account)
            grouped = group.groupby('Po Number')
            group = grouped.agg(col_methods)
            row_sum = group[row_sum_names].sum(axis=1)
            group['Amount'] = row_sum
            col_sum = group[col_sum_names].sum()
            col_sum = pd.Series(col_sum.values, index=[['Total', 'Total', 'Total', 'Total', 'Total', 'Total', 'Total'], col_sum_names])
            col_sum = col_sum.unstack()
            group = pd.concat([group, col_sum])
            group = group.reindex(columns=out_col_filters)
            group.drop('Entity',axis=1, inplace=True)
            group.drop('Account',axis=1, inplace=True)
            group['Remark'] = ""
            group.rename(columns = out_col_names, inplace=True)
            group.to_excel(writer, sheet_name=name)
        try:
            writer.save()
        except Exception as e:
            self.context_text.insert(self.context_text_index, e)
            self.context_text_index += 1
            return '请重新打开文件！！！'
        self.combine_sheets(out_file_name)
        return out_file_name

    def combine_sheets(self,in_file_name):
        out_col_filters = ['Po Number', 'Vendor Code', 'Supplier',
                           '0-30 Days CNY', '31-60 Days CNY', '61-90 Days CNY', '91-180 Days CNY',
                           '181-360 Days CNY', 'Over 360 Days CNY', 'Amount', 'Remark']
        group = pd.DataFrame()
        excel_name = in_file_name.rstrip('.xls') +"_combine.xlsx"
        writer = pd.ExcelWriter(excel_name)
        xls_file = pd.ExcelFile(in_file_name)
        b = xlrd.open_workbook(in_file_name)
        entity = ""
        before_entity = ""
        begin_flag = True
        for sheet in b.sheets():
            try:
                sheet_name_arr = sheet.name.split(" ")
                before_entity = entity
                entity = sheet_name_arr[0]
                account = sheet_name_arr[1]
                if(before_entity == ""  and begin_flag == True):
                    time = self.getLastDayOfLastMonth()
                    entity_row = pd.Series([' ', "GRIR Aging-"+entity, ' ', ' ', ' ',' ', ' ',' ', ' ',' ',' '],
                                         index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], out_col_filters])
                    entity_row = entity_row.unstack()
                    time_row = pd.Series([' ', time, ' ', ' ', ' ',' ', ' ',' ', ' ',' ',' '],
                                         index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], out_col_filters])
                    time_row = time_row.unstack()
                    group = pd.concat([entity_row, time_row, group], axis=0)
                if before_entity != entity and begin_flag != True:
                    begin_flag = True
                    group = group.reindex(columns=out_col_filters)
                    entity_name = self.convert_entity_name(before_entity)
                    group.to_excel(writer, sheet_name=entity_name)
                    group = pd.DataFrame()
                    now = datetime.now().strftime('%b,%d,%Y')
                    entity_row = pd.Series([' ', "GRIR Aging-"+entity, ' ', ' ', ' ',' ', ' ',' ', ' ',' ',' '],
                                         index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], data.columns])
                    entity_row = entity_row.unstack()
                    time_row = pd.Series([' ', time, ' ', ' ', ' ',' ', ' ',' ', ' ',' ',' '],
                                         index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], data.columns])
                    time_row = time_row.unstack()
                    group = pd.concat([entity_row, time_row, group], axis=0)
                begin_flag = False
                data = xls_file.parse(sheet.name, header=0)
                blank_row = pd.Series("",index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], data.columns])
                blank_row = blank_row.unstack()
                columns_row = pd.Series(data.columns,index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], data.columns])
                columns_row = columns_row.unstack()
                name_row = pd.Series([entity, account, ' ', ' ', ' ',' ', ' ',' ', ' ',' ',' '],
                                     index=[[' ', ' ', ' ',' ', ' ',' ', ' ',' ', ' ',' ',' '], data.columns])
                name_row = name_row.unstack()
                data = pd.concat([name_row, columns_row, data], axis=0)
                group = pd.concat([group, blank_row, blank_row, blank_row, blank_row, data], axis=0)
            except Exception as e:
                self.context_text.insert(self.context_text_index, excel_name+'生成错误\n\r')
                self.context_text_index += 1
                return '请重新打开文件！！！'
        group = group.reindex(columns=out_col_filters)
        entity_name = self.convert_entity_name(entity)
        group.to_excel(writer, sheet_name=entity_name)
        group = pd.DataFrame()     
        try:
            writer.save()
        except Exception as e:
            self.context_text.insert(self.context_text_index, e)
            self.context_text_index += 1
            return '请重新打开文件！！！'

    def convert_entity_name(self,entity):
        name = ""
        if entity == "982":
            name = "SH HQ"
        elif entity == "983":
            name = "SH CCS"
        elif entity == "985":
            name = "SH Plant"
        elif entity == "1500":
            name = "BZ plant"
        elif entity == "1520":
            name = "CX Plant"
        elif entity == "1530":
            name = "SY Plant"
        elif entity == "1550":
            name = "CQ plant"
        elif entity == "1570":
            name = "JY Plant"
        elif entity == "1990":
            name = "APS"
        return name


    def getLastDayOfLastMonth(self):
        d = datetime.now()
        year = d.year
        month = d.month
        if month == 1:
            month = 12
            year -= 1
        else:
            month -= 1
        days = calendar.monthrange(year, month)[1]
        return (datetime(year, month, 1) + timedelta(days=days-1)).strftime('%b,%d,%Y')

    def get_desktop(self):
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,\
                              r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
        return winreg.QueryValueEx(key, "Desktop")[0]

    def mkdir(self,path):
        # 去除首位空格
        path=path.strip()
        # 去除尾部 \ 符号
        path=path.rstrip("\\")
        isExists=os.path.exists(path)
        # 判断结果
        if not isExists:
            # 创建目录操作函数
            os.makedirs(path)
            print(path+' 创建成功')
            return True
        else:
            # 如果目录存在则不创建，并提示目录已存在
            print(path+' 目录已存在')
            return False

    def openfile(self):
        # 定义要创建的目录
        #kpath=os.path.abspath('..') + '\DATA'
        mkpath=self.get_desktop() + '\DATA'
        # 调用函数
        self.mkdir(mkpath)
        in_file_name = filedialog.askopenfilename(title='Open File', filetypes=[('excel', '*.xls *.xlsx'), ('All Files', '*')])
        out_file_name = self.processexcel(in_file_name)
        if out_file_name == "请重新打开文件！！！":
            self.context_text.insert(self.context_text_index, out_file_name + '\n')
            self.context_text_index += 1
            return
        self.context_text.insert(self.context_text_index, '成功生成文件：' + out_file_name + '\n')
        self.context_text_index += 1


    def close(self):
        self.root.destroy()
