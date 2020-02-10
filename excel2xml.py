# -*- coding:utf-8 -*-
import os
import sys
# reload(sys)
# sys.setdefaultencoding( "utf-8" )

from excelParse import ExcelParser
class operate():
    def __init__(self, ExcelFileName,SheetName):
        self.excelFile = ExcelFileName
        self.SheetName = SheetName
        self.temp = ExcelParser(self.excelFile)
        self.dic_testlink = {}
        self.row_flag = 3
        self.testsuite = SheetName
        self.dic_testlink[self.testsuite] = {"node_order": "13", "details": "", "testcase": []}
        self.content = ""
        self.content_list = []
    def excel2dic(self,SheetName):
        # excel数据转换为字典
        # 获取首行表名称
        column_name_list=self.temp.get_one_row_content(0,SheetName)
        custom_list_all=column_name_list.copy()
        custom_list=[]
        for custom_filed in custom_list_all:
            if "$" in custom_filed:
                # print(custom_filed)
                custom_list.append(custom_filed)

        colum_index=column_name_list.index("name")
        testcase_list=self.temp.get_one_colum_content(colum_index, SheetName)
        #print testcase_list
        #获取每个testcase的行号，从而得出step步骤N行  testcase_row_num_list 用例的行号
        testcase_row_num_list=[]
        for i in range(1, len(testcase_list)):
            if testcase_list[i]!="":
                testcase_row_num_list.append(i)
        
        #print testcase_row_num_list
        steps_rows=self.temp.get_one_colum_content(4, SheetName)
        rows_sum=len(steps_rows)
        #获取每个case各个属性值组成每个testcase的map
        for row_num in testcase_row_num_list:
            testcase = {"name": "", "node_order": "100", "version": "1", "summary": "",
                        "preconditions": "", "execution_type": "1", "importance": "", "steps": [], "keywords": "",
                        "custom_fields":[]}
            case_name=self.temp.get_one_cell_content(row_num,"name",SheetName)
            summary=self.temp.get_one_cell_content(row_num,"summary",SheetName)
            preconditions=self.temp.get_one_cell_content(row_num,"preconditions",SheetName)
            importance=self.temp.get_one_cell_content(row_num,"importance",SheetName)
            keywords=self.temp.get_one_cell_content(row_num,"keywords",SheetName)
            execution_type=self.temp.get_one_cell_content(row_num,"execution_type",SheetName)
            # 自定义字段添加
            for custom_filed in custom_list:
                # print(custom_filed)
                tmp=self.temp.get_one_cell_content(row_num,custom_filed,SheetName)
                testcase["custom_fields"].append({custom_filed:tmp})

            # reverse=self.temp.get_one_cell_content(row_num,"reverse",SheetName)
            num_index=testcase_row_num_list.index(row_num)
            #print num_index
            #print len(testcase_row_num_list)
            # 处理跨行多个步骤
            if num_index!=len(testcase_row_num_list)-1:
                next_row_num=testcase_row_num_list[num_index+1]
                step_num=1
                for i in range(row_num, next_row_num):
                    step= {"step_number": "", "actions": "", "expectedresults": "", "execution_type": ""}
                    action=self.temp.get_one_cell_content(i,"actions",SheetName)
                    expect_result=self.temp.get_one_cell_content(i,"expected_result",SheetName)
                    execution_type=self.temp.get_one_cell_content(i,"execution_type",SheetName)
                    step["step_number"]=str(step_num)
                    step["actions"]=action
                    step["expectedresults"]=expect_result
                    step["execution_type"]=execution_type
                    step_num+=1      
                    testcase["steps"].append(step)
            else:
                step_num=1
                for i in range(row_num, rows_sum):
                    step= {"step_number": "", "actions": "", "expectedresults": "", "execution_type": ""}
                    action=self.temp.get_one_cell_content(i,"actions",SheetName)
                    expect_result=self.temp.get_one_cell_content(i,"expected_result",SheetName)
                    execution_type=self.temp.get_one_cell_content(i,"execution_type",SheetName)
                    step["step_number"]=str(step_num)
                    step["actions"]=action
                    step["expectedresults"]=expect_result
                    step["execution_type"]=execution_type
                    step_num+=1
                    testcase["steps"].append(step)

            testcase["name"]=case_name
            testcase["summary"]=summary
            testcase["preconditions"]=preconditions
            testcase["importance"]=importance
            testcase["keywords"]=keywords
            # 将自定义字段添加
            # testcase["is_reverse"]=is_reverse
            #testcase["steps"].append(step)
            # for colume_name in all_case_name:
            #     if colume_name =="custom":
            #     custom_field={"name":"","value":""}
            #     testcase["custom_fields"].append(custom_field)
            testcase["execution_type"]=execution_type
            self.dic_testlink[self.testsuite]["testcase"].append(testcase)

    def content_to_xml(self, key, value=None):
        if key == 'step_number' or  key == 'node_order' or key == 'version' :
            return "<" + str(key) + "><![CDATA[" + str(value) + "]]></" + str(key) + ">\n"
        # 自定义字段
        elif key=='custom_fields':
            results=""
            for custom_key in value:
                (tkey,tvalue),=custom_key.items()
                tkey=tkey.replace("$","")
                result = f"<custom_field><name><![CDATA[{tkey}]]></name><value><![CDATA[{tvalue}]]></value></custom_field>\n"
                results=results+result
            return "<custom_fields>"+results+"</custom_fields>\n"
        elif key == 'execution_type':
            if value == "自动":
                convert_value = "2"
            else :
                convert_value = "1"
            # print(value,convert_value)
            return "<" + str(key) + "><![CDATA[" + str(convert_value) + "]]></" + str(key) + ">\n"
        elif  key == 'importance':
            if value == "高":
                convert_value = "3"
            elif value=="中" :
                convert_value = "2"
            else:
                convert_value = "1"
            # print(value,convert_value)
            return "<" + str(key) + "><![CDATA[" + str(convert_value) + "]]></" + str(key) + ">\n"
        elif key == 'actions' or key == 'expectedresults' or key == 'summary' or key == 'preconditions':
            # 多行需要换行问
            outvalue=""
            p_list=value.split('\n')
            for single_p in p_list:
                line=f"<p>{single_p}</p>"
                outvalue=outvalue+line
            return "<" + str(key) + "><![CDATA[<p> " + str(outvalue) + "</p> ]]></" + str(key) + ">\n"
        elif key == 'keywords':
            if "," in value:
                content_pre='<keywords>\n'
                content_end='</keywords>\n'
                content_mid=''
                value_list=value.split(',')
                for value in value_list:
                    content_mid=content_mid+'<keyword name="' + str(value) + '">\n<notes><![CDATA[ keyowrd ]]></notes>\n</keyword>\n'
                return content_pre+content_mid+content_end
            else:
                return '<keywords>\n<keyword name="' + str(value) + '">\n<notes><![CDATA[ keyowrd ]]></notes>\n</keyword>\n</keywords>\n'
        elif key == 'name':
            return '<testcase name="' + str(value) + '">\n'
        else:
            return '*ERROR*'

    def dic_to_xml(self):
        # 将一个表的数据转换为xml数据
        self.excel2dic(self.SheetName)
        testcase_list = self.dic_testlink[self.testsuite]["testcase"]
        for testcase in testcase_list:
            for step in testcase["steps"]:
                self.content += "<step>\n"
                self.content += self.content_to_xml("step_number", step["step_number"])
                self.content += self.content_to_xml("actions", step["actions"])
                self.content += self.content_to_xml("expectedresults", step["expectedresults"])
                self.content += self.content_to_xml("execution_type", step["execution_type"])
                self.content += "</step>\n"
            self.content = "<steps>\n" + self.content + "</steps>\n"
            self.content = self.content_to_xml("importance", testcase["importance"]) + self.content
            self.content = self.content_to_xml("execution_type", testcase["execution_type"]) + self.content
            self.content = self.content_to_xml("preconditions", testcase["preconditions"]) + self.content
            self.content = self.content_to_xml("summary", testcase["summary"]) + self.content
            self.content = self.content_to_xml("version", testcase["version"]) + self.content
            #self.content = self.content_to_xml("externalid", testcase["externalid"]) + self.content
            self.content = self.content_to_xml("node_order", testcase["node_order"]) + self.content
            self.content = self.content + self.content_to_xml("keywords", testcase["keywords"])
            # 自定义字段添加
            self.content = self.content+self.content_to_xml("custom_fields",testcase["custom_fields"])
            self.content = self.content_to_xml("name", testcase["name"]) + self.content
            self.content = self.content + "</testcase>\n"
            self.content_list.append(self.content)
            self.content = ""
        self.content = "".join(self.content_list)
        self.content = '<testsuite name="' + self.testsuite + '">\n' + self.content + "</testsuite>"
        self.content = '<?xml version="1.0" encoding="UTF-8"?>\n' + self.content
        self.write_to_file(self.excelFile)

    # 输出xml文件
    def write_to_file(self, ExcelFileName):
        xmlFileName = ExcelFileName.split('.')[0] + '_' + self.SheetName + '.xml'
        cp = open(xmlFileName, "w")
        cp.write(self.content)
        cp.close()

if __name__ == "__main__":
    op=operate("点亮计划.xlsx","黑名单")
    op.dic_to_xml()    
