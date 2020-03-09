# testlink_excel2xml
将excel文件转换为xml文件导入到testlink
使用注意事项：
1.自定义字段使用$符号为前缀。如 $结果
2.特殊符号在模板中支持有限，如> < 和html标签有冲突，会导致导入模版报错。目前只有name字段能够兼容特殊符号。

使用教程：
1.使用Main_excel2xml.py，修改file参数文件路径。
