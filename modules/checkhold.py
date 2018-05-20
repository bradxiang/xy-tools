# -*- coding: UTF-8 -*-
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
from datetime import datetime
from datetime import timedelta
import os
import winreg


class CheckHold():
    def __init__(self):
        self.EPSINON = 0.00001
        self.args = None
        self.supplier_list = ['ILHCJS', 'ILDPJC', 'ILZJHJ', 'ILYMSC', 'BXZJHB', 'ILHKGG',
                              'ILLNHG', 'ILDBYJ', 'ILYHSH', 'ILHJBH', 'ILSWHK', 'ILSJJT',
                              'BXCQGH', 'ILLNAK', 'ILSYGJ', 'ILHZZR', 'ILBTNS', 'ILSDHK',
                              'ILHNBP', 'ILHNYJ', 'ILZJYJ']
        self.supplier_names = ""
        for index, supplier_name in enumerate(self.supplier_list):
            self.supplier_names += str(index) + ')' + supplier_name + ' '
        self.contents = None

    def processexcel(self, in_file_name):
        sheet_name = "Sheet1"
        temp = in_file_name.split('\\')
        out_file_name = self.get_desktop() + '\DATA\CheckHold-output-' + temp[len(temp) - 1]
        if self.args['AutoHold_var'] == "" or (not self.args['AutoHold_var'].isdigit()):
            self.contents.write('错误:AutoHoldDays为空或不合法请输入1-99！！！\n')
            return "请重新打开文件！！！"
        if self.args['TC_HOLD_AMOUNT_var'] == "" or (not self.args['TC_HOLD_AMOUNT_var'].isdigit()):
            self.contents.write('TC_HOLD_AMOUNTDays为空或不合法请输入1-99！！！\n')
            return '请重新打开文件！！！'
        self.AutoHold_due = int(self.args['AutoHold_var'])
        self.TC_HOLD_AMOUNT_due = int(self.args['TC_HOLD_AMOUNT_var'])
        in_col_filters = ['Entity', 'Supp/Cust Code', 'Business Relation Name1', 'XX_POSO', 'Control GL Account',
                          'Reference', 'XX_INVOICE_S',
                          'XX_INVOICE_DESC', 'Invoice Type', 'INV_CURRENCY', 'TC Balance', 'Due Date', 'TC_HOLD_AMOUNT']
        out_col_names = {}
        xls_file = pd.ExcelFile(in_file_name)
        try:
            raw = xls_file.parse(sheet_name, fill_value=0)
        except Exception as e:
            self.contents.write('文件选择错误或sheet重命名为Sheet1！！！\n')
            return '请重新打开文件！！！'
        try:
            table = raw[in_col_filters]
        except Exception as e:
            self.contents.write('文件头数据不匹配，请检查！！！\n')
            return '请重新打开文件！！！'
        writer = pd.ExcelWriter(out_file_name)
        for XX_INVOICE_S, group in table.groupby('XX_INVOICE_S'):
            if XX_INVOICE_S == 'Auto-Hold':
                name = XX_INVOICE_S
                name += '<=' + str(self.AutoHold_due)
                group1 = group
        supplier_group = self.get_supplier_data(table, self.supplier_list)
        group1 = pd.concat([group1, supplier_group], axis=0)
        due_time = pd.to_datetime(group1['Due Date'])
        time_select = due_time <= (datetime.now() + timedelta(self.AutoHold_due))
        group1 = group1[time_select]

        group1 = group1[~group1.index.duplicated()]

        group1.to_excel(writer, sheet_name=name)
        name1 = 'TC_HOLD_AMOUNT<=' + str(self.TC_HOLD_AMOUNT_due)
        due_time = pd.to_datetime(table['Due Date'])
        time_select = due_time <= (datetime.now() + timedelta(self.TC_HOLD_AMOUNT_due))
        group2 = table[time_select]
        num_select = group2['TC_HOLD_AMOUNT'] > 0
        group2 = group2[num_select]

        group2 = group2[~group2.index.duplicated()]

        group2.to_excel(writer, sheet_name=name1)
        try:
            writer.save()
        except Exception as e:
            self.contents.write(e)
            return '请重新打开文件！！！'
        return out_file_name

    def get_supplier_data(self, table, supplier_list):
        supplier_group = pd.DataFrame()
        for supplier_name in supplier_list:
            supplier_instance = table[table['Supp/Cust Code'] == supplier_name]
            supplier_group = pd.concat([supplier_group, supplier_instance], axis=0)
        return supplier_group

    def get_desktop(self):
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\Current' +
            r'Version\Explorer\Shell Folders',
        )
        return winreg.QueryValueEx(key, "Desktop")[0]

    def mkdir(self, path):
        # 去除首位空格
        path = path.strip()
        # 去除尾部 \ 符号
        path = path.rstrip("\\")
        isExists = os.path.exists(path)
        # 判断结果
        if not isExists:
            # 创建目录操作函数
            os.makedirs(path)
            print(path + ' 创建成功')
            return True
        else:
            # 如果目录存在则不创建，并提示目录已存在
            print(path + ' 目录已存在')
            return False

    def run(self, args):
        self.contents = args['contents']
        in_file_name = args['in_file_name']
        self.contents.write('Supp/Cust Code includes: ' + self.supplier_names + '\n')
        self.args = args
        mkpath = self.get_desktop() + '\DATA'
        # 调用函数
        self.mkdir(mkpath)
        out_file_name = self.processexcel(in_file_name)
        if out_file_name == "请重新打开文件！！！":
            self.contents.write(out_file_name + '\n')
            return
        self.contents.write('成功生成文件：' + out_file_name + '\n')
