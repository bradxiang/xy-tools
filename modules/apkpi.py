# -*- coding: UTF-8 -*-
import os
import winreg
import re
import math
import xlrd
import pandas as pd


class ApKpi():
    def __init__(self):
        pass

    def process_excel(self, in_file_name, contents):
        data_columns = None
        group = pd.DataFrame()
        temp = in_file_name.split('\\')
        out_file_name = self.get_desktop(
        ) + '\\DATA\\kpiapaging-output-' + temp[len(temp) - 1]
        xls_file = pd.ExcelFile(in_file_name)
        b = xlrd.open_workbook(in_file_name)
        contents.write(
            'entity一共处理9个:0982、0983、0985、1500、1520、1530、1550、1570、1990！！！\n')
        for sheet in b.sheets():
            if sheet.name.isdigit():
                try:
                    data = xls_file.parse(sheet.name, fill_value=0)
                except Exception as e:
                    contents.write('ERROR: ' + e + '，文件无法处理，请上传正确格式文件！！！\n')
                    return
                data.insert(0, "entity", sheet.name)
                data.insert(1, "plant", sheet.name)
                if sheet.name == "0982":
                    data.loc[:, 'plant'] = "SH HQ"
                elif sheet.name == "0983":
                    data.loc[:, 'plant'] = "SH CCS"
                elif sheet.name == "0985":
                    data.loc[:, 'plant'] = "SH Plant"
                elif sheet.name == "1500":
                    data.loc[:, 'plant'] = "BZ plant"
                elif sheet.name == "1520":
                    data.loc[:, 'plant'] = "CX Plant"
                elif sheet.name == "1530":
                    data.loc[:, 'plant'] = "SY Plant"
                elif sheet.name == "1550":
                    data.loc[:, 'plant'] = "CQ plant"
                elif sheet.name == "1570":
                    data.loc[:, 'plant'] = "JY Plant"
                elif sheet.name == "1990":
                    data.loc[:, 'plant'] = "APS"
                else:
                    contents.write('特殊plant:' + sheet.name + '请检查数据！！！\n')
                    continue
                del_index = []
                for i in range(data.shape[0]):
                    BRM = data.loc[i, "Supplier"]
                    if BRM == 'Summaries:':
                        del_index.append(i)
                    elif type(BRM) == float:
                        if math.isnan(BRM):
                            del_index.append(i)
                data = data.drop(data.index[del_index])
                data.insert(2, "IC/3RD", "")
                for i in range(data.shape[0]):
                    ic3rd_flag = False
                    BRM = data.loc[i, "Business Relation Name"]
                    s = re.match(r'(.*)江森(.*)', BRM)
                    if str(s) != 'None':
                        if BRM != "重庆博奥江森蓄电池有限公司":
                            data.loc[i, "IC/3RD"] = "IC"
                            ic3rd_flag = True
                    s = re.match(r'(.*)Johnson Controls(.*)', BRM)
                    if str(s) != 'None':
                        data.loc[i, "IC/3RD"] = "IC"
                        ic3rd_flag = True
                    s = re.match(r'(.*)约克', BRM)
                    if str(s) != 'None':
                        data.loc[i, "IC/3RD"] = "IC"
                        ic3rd_flag = True
                    if BRM == 'ENERTEC EXPORTS SDE RL DE CV':
                        data.loc[i, "IC/3RD"] = "IC"
                        ic3rd_flag = True
                    rf = str(data.loc[i, "Reference"])
                    s = re.match(r'(.*)ER(.*)', rf)
                    if str(s) != 'None':
                        data.loc[i, "IC/3RD"] = "TE"
                        ic3rd_flag = True
                    isc = data.loc[i, "Invoice Status Code"]
                    if isc == 'GATES':
                        data.loc[i, "IC/3RD"] = "TE"
                        ic3rd_flag = True
                    if ic3rd_flag is False:
                        data.loc[i, "IC/3RD"] = "3rd"
                data.insert(3, "nonpo/po", "")
                for i in range(data.shape[0]):
                    nonpo_flag = False
                    tp = data.loc[i, "Type"]
                    if tp == 'Invoice Correction':
                        data.loc[i, "nonpo/po"] = "Invoice Correction"
                        nonpo_flag = True
                    tp = data.loc[i, "Type"]
                    if tp == 'Credit Note Correction':
                        data.loc[i, "nonpo/po"] = "Credit Note Correction"
                        nonpo_flag = True
                    tp = data.loc[i, "Type"]
                    if tp == 'Prepayment':
                        data.loc[i, "nonpo/po"] = "Prepayment"
                        nonpo_flag = True
                    fpo = data.loc[i, "First PO Number"]
                    if type(fpo) != float:
                        if nonpo_flag is False:
                            data.loc[i, "nonpo/po"] = "PO"
                            nonpo_flag = True
                    if nonpo_flag is False:
                        data.loc[i, "nonpo/po"] = "NON-PO"
                    if data_columns is None:
                        data_columns = data.columns
                group = pd.concat([group, data], axis=0)
                contents.write('完成entity: ' + sheet.name + '\n')
            else:
                contents.write('文件选择错误或将待处理表重命名为数字，若'+ sheet.name + '表无用请忽略本条信息！！！\n')
        group = group[data_columns.values]
        writer = pd.ExcelWriter(out_file_name)
        name = "KPI-AP-AGING"
        group.to_excel(writer, sheet_name=name)
        try:
            writer.save()
            contents.write('生成文件' + out_file_name + '\n')
        except Exception as e:
            contents.write(e + '文件生成错误\n')

    @staticmethod
    def get_desktop():
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]

    @staticmethod
    def create_directory(path, contents):
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

    def run(self, in_file_name, contents):
        contents.write('******************开始处理*******************\n')
        path = self.get_desktop() + '\\DATA'  # 定义要创建的目录
        self.create_directory(path, contents)
        self.process_excel(in_file_name, contents)
        contents.write('******************结束处理*******************\n')
