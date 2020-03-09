# -*- coding: utf-8 -*-
"""
@File Name：Main_excel2xml
@Software: PyCharm
@Author : MaoXianfei
@Time：2020/2/10 12:13 上午 
"""
from excelParse import ExcelParser
from excel2xml import operate

def Control(file):
    sheetnames=ExcelParser(file).get_sheet_name()
    for sheetname in sheetnames:
        print(f"{sheetname} - 模块完成")
        op = operate(file, sheetname)
        op.dic_to_xml()
if __name__ == '__main__':
    # 一个sheet表输出一个xml文件
    file="fccc.xlsx"
    Control(file)