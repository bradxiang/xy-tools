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
            self.supplier_names += str(index) + ')' + supplier_name + '  '

    def processexcel(self, in_file_name, contents):
        sheet_name = "Sheet1"
        temp = in_file_name.split('\\')
        out_file_name = self.get_desktop() + '\\DATA\\CheckHold-output-' + temp[len(temp) - 1]
        if self.args['AutoHold_var'] == "" or (not self.args['AutoHold_var'].isdigit()):
            contents.write('错误:AutoHoldDays为空或不合法请输入1-99！！！\n')
            return False
        if self.args['TC_HOLD_AMOUNT_var'] == "" or (not self.args['TC_HOLD_AMOUNT_var'].isdigit()):
            contents.write('TC_HOLD_AMOUNTDays为空或不合法请输入1-99！！！\n')
            return False
        self.AutoHold_due = int(self.args['AutoHold_var'])
        self.TC_HOLD_AMOUNT_due = int(self.args['TC_HOLD_AMOUNT_var'])
        in_col_filters = ['Entity', 'Supp/Cust Code', 'Business Relation Name1', 'XX_POSO', 'Control GL Account',
                          'Reference', 'XX_INVOICE_S', 'XX_INVOICE_DESC', 'Invoice Type', 'INV_CURRENCY', 'TC Balance',
                          'Due Date', 'TC_HOLD_AMOUNT']
        out_col_names = {}
        xls_file = pd.ExcelFile(in_file_name)
        try:
            raw = xls_file.parse(sheet_name, fill_value=0)
        except Exception as e:
            contents.write('文件选择错误或将待处理表重命名为Sheet1！！！\n')
            return False
        try:
            table = raw[in_col_filters]
        except Exception as e:
            contents.write('文件列数据不匹配，请检查！！！\n')
            return False
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
            contents.write('输出文件数据写入失败！！！\n')
            contents.write(e)
            return False
        contents.write('成功生成文件：' + out_file_name + '\n')
        return True

    def get_supplier_data(self, table, supplier_list):
        supplier_group = pd.DataFrame()
        for supplier_name in supplier_list:
            supplier_instance = table[table['Supp/Cust Code'] == supplier_name]
            supplier_group = pd.concat([supplier_group, supplier_instance], axis=0)
        return supplier_group

    def get_desktop(self):
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]

    def create_directory(self, path, contents):
        # 去除首位空格
        path = path.strip()
        # 去除尾部 \ 符号
        path = path.rstrip("\\")
        if not os.path.exists(path):
            try:
                os.makedirs(path)
                contents.write(path + '创建成功\n')
                return True
            except Exception as e:
                contents.write('ERROR: ' + e + '\n请手动在桌面创建名称为DATA文件夹\n')
                return False
        else:
            contents.write(path + '目录已存在\n')
            return False

    def run(self, args):
        contents = args['contents']
        in_file_name = args['in_file_name']
        contents.write('******************开始处理*******************\n')
        contents.write('Supp/Cust Code includes: ' + self.supplier_names + '\n')
        self.args = args
        desktop_path = self.get_desktop() + '\\DATA'
        self.create_directory(desktop_path, contents)
        out_file_name = self.processexcel(in_file_name, contents)
        if out_file_name == False:
            contents.write('CheckHold处理失败！！！\n')
        else:
            contents.write('CheckHold处理成功！！！\n')
        contents.write('******************结束处理*******************\n')
