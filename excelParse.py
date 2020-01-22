#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
Class that holds information about parsing Excel file.
"""

import xlrd
import xlwt    
import os,sys
from xml.etree import ElementTree
# reload(sys)
# sys.getdefaultencoding()
# sys.setdefaultencoding( "utf-8" )


class ExcelParser :
    """
    Parse Excel file.
    """ 
    def __init__(self, path) :
        self.file_path = None
        self.book = None
        self.excel_sh = None
        self.current_sheet_name = None
        self.write_book = None
        self.file_path = path
        if os.path.exists(self.file_path) is False:
            raise AssertionError("File not existed:" + path)
        try:
            self.book = xlrd.open_workbook(self.file_path)
        except Exception as e:
            print("Excel isn't well formated:" + e.message)
            raise
      
    def get_one_row_content(self, row, sheetname = None) :
        """
        get data by row
        - @var row: The row number of the excel
        - @var sheetname: The default sheetname is None, 
                    and we will change the last time's sheet name. 
        """
        if sheetname != None:
            self.current_sheet_name = sheetname
        excel_sh = self.book.sheet_by_name(self.current_sheet_name) 
        return excel_sh.row_values(row)
    
    def get_one_colum_content(self, column, sheetname = None) :
        """
        get data by column
        - @var column: The column number of the excel 
        - @var sheetname: The default sheetname is None, 
                    and we will change the last time's sheetname.
        """
        if sheetname != None:
            self.current_sheet_name = sheetname
        excel_sh = self.book.sheet_by_name(self.current_sheet_name)
        return excel_sh.col_values(column)
    
    def get_file_rows(self, sheetname = None) :
        """
        get current sheet's total rows
         - @var sheetname: The default sheetname is None, 
                    and we will change the last time's sheetname.
        """
        if sheetname != None:
            self.current_sheet_name = sheetname
        excel_sh = self.book.sheet_by_name(self.current_sheet_name)
        return excel_sh.nrows
    
    def get_file_colum(self, sheetname = None) :
        """
        get current sheet's total column
         - @var index: The default sheetname is None, 
                    and we will change the last time's sheetname.
        """ 
        if sheetname != None:
            self.current_sheet_name = sheetname
        excel_sh = self.book.sheet_by_name(self.current_sheet_name)
        return excel_sh.ncols
            
    def get_one_cell_content(self, row, column_name, sheetname = None) :
        """
        get special cell's data
        - @var column_name: The column name of the excel 
        - @var row: The row number of the excel 
        - @var sheetname: The default sheetname is None, 
                    and we will change the last time's sheetname.
        """
        if sheetname != None:
            self.current_sheet_name = sheetname
        excel_sh = self.book.sheet_by_name(self.current_sheet_name)
        column_name_list=self.get_one_row_content(0,self.current_sheet_name)
        colum_index=column_name_list.index(column_name)
        #print colum_index
        return excel_sh.cell_value(row, colum_index)
    
    def save(self, file_name = "test.xls"):   
        """
        save the excel file to the specialized file name 
        - @var file_name: File name of the file you want to save as.
        """
        self.write_book.save(file_name)
        
    def set_value(self, value, row, column, sheetname = None):
        """
        Set a special cell of value 
        - @var value: The value which you want to set in the cell.
        - @var column: The column number of the excel 
        - @var row: The row number of the excel 
        - @var index: The default index number is None, 
                    and we will change the last time's sheet number.
        """
        if sheetname != None:
            self.current_sheet_name = sheetname
        write_sheet = self.write_book.get_sheet(self.current_sheet_name)
        write_sheet._cell_overwrite_ok = True 
        write_sheet.write(row, column, value)
        
if __name__ == "__main__":
    filename = "做市商借款测试计划_v1.0_20191230.xlsx"
    exparse=ExcelParser(filename)
    testcase_list=exparse.get_one_colum_content(0,"测试用例")
    print(testcase_list)

    