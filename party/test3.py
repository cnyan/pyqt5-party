# !/usr/bin/python3
# -*- coding: utf-8 -*-

"""
@Author: 闫继龙
@Version: ??
@License: Apache Licence
@CreateTime: 2019/9/9 16:05
@Describe：

"""
# !/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2018/6/13 0013 8:42
# @Author  : 一梦南柯
# @File    : upload_data4.py

'''
用python读取xml格式的excel文件
'''
# coding=utf-8
import xml
from xml.dom import minidom
import re


class get_xml():
    # 加载获取xml的文档对象
    def __init__(self, address):
        # 解析address文件，返回DOM对象，address为文件地址
        self.doc = minidom.parse(address)
        # DOM是树形结构，_get_documentElement()获得了树形结构的根节点
        self.root = self.doc._get_documentElement()
        # .getElementsByTagName()，根据name查找根目录下的子节点
        self.httpSample_nodes = self.root.getElementsByTagName('httpSample')

    def getxmldata(self):
        data_list = []
        j = -1
        responseData_node = self.root.getElementsByTagName("responseData")
        for i in self.httpSample_nodes:
            j = j + 1
            # getAttribute()，获取DOM节点的属性的值
            if i.getAttribute("lb") == "发送信息":
                a = 'chatId":"(.*?)"'
            elif i.getAttribute("lb") == "接收信息":
                # a = "chatId%3A%22(.*?)%22"
                a = "info%3A%22(.*?)%22"
            if (i.getAttribute("lb") == "发送信息" or i.getAttribute("lb") == "接收信息") and i.getAttribute("rc") == "200":
                try:
                    # 使用re包里面的方法，通过正则表达式提取数据
                    b = re.search(a, responseData_node[j].firstChild.data)
                    if b is not None:
                        d = b.group(1)
                        print("d:", d)
                        data_list.append((d, i.getAttribute("ts"), i.getAttribute("lt"), i.getAttribute("lb")))
                except:
                    pass
        return data_list


if __name__=='__main__':
    file_addr = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls'
    x = get_xml(file_addr)
    print(x)