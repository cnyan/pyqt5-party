# !/usr/bin/python3
# -*- coding: utf-8 -*-

"""
@Author: 闫继龙
@Version: ??
@License: Apache Licence
@CreateTime: 2019/8/26 10:28
@Describe：

"""

import sys
from PyQt5.QtWidgets import (QWidget, QGridLayout, QLabel, QLineEdit,
                             QTextEdit, QPushButton, QApplication, QDesktopWidget)
import xlrd  # 读取excel
import xlwt  # 写入excel
import numpy as np
import os
import json


def all_path(dirname):
    result = []
    for maindir, subdir, file_name_list in os.walk(dirname):
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)
            result.append(apath)
    # print(result)
    # 列表排序
    result.sort(key=lambda x: int(x.split('\\')[-1].split('.')[0]))
    return result  # 返回文件列表

class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.fileDir = 'D:\home\partydb'
        self.initUI()

    def initUI(self):
        self.setWindowTitle("party")
        title = QLabel('fileDir')
        self.titleEdit = QLineEdit('')
        titleBtn = QPushButton('submit', self)
        titleBtn.clicked.connect(self.buttonClicked)

        # 处理btn
        work1Btn = QPushButton('work1', self)
        work1Btn.clicked.connect(self.work1)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(title, 1, 0)
        grid.addWidget(self.titleEdit, 1, 1)
        grid.addWidget(titleBtn, 1, 2)
        grid.addWidget(work1Btn, 2, 1)

        self.setLayout(grid)
        self.center()
        self.setGeometry(300, 300, 350, 300)
        self.show()

    # 控制窗口显示在屏幕中心的方法
    def center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def buttonClicked(self):
        sender = self.sender()

        fileEdit = self.titleEdit.text()
        if fileEdit != '':
            self.fileDir = fileEdit

        print(self.fileDir)
        print(sender.text())
        # self.statusBar().showMessage(sender.text() + ' was pressed')

    def work1(self):
        sender = self.sender()

        rootdir = os.path.join(self.fileDir, '表1')
        excel_list = all_path(rootdir)
        print(excel_list)
        row_dict = {}

        # 所有数据
        work_data = np.zeros((3, 12), dtype='int16')

        # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
        for file_excel in excel_list:
            by_sheet = u'附件3—党组织情况统计表（表1）'
            # 打开数据所在的工作簿，以及选择存有数据的工作表
            book = xlrd.open_workbook(file_excel)
            sheet = book.sheet_by_name(by_sheet)
            n_rows = sheet.nrows  # 行数

            sheet_data = []  # 表数据
            # data_sheet = np.zeros((3, 12))
            # 按行遍历一张表
            for row_num in range(6, 9):
                row = sheet.row_values(row_num)
                row_dict[row_num] = row

                row_data = []  # 行数据
                col_list = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
                for c in range(len(col_list)):
                    df = row[col_list[c]]
                    if isinstance(df, str):
                        df = 0
                    else:
                        df = int(df)
                    row_data.append(df)
                sheet_data.append(row_data)

            sheet_data = np.array(sheet_data)
            # print(sheet_data)
            work_data += sheet_data
            # print('88888888888888' * 8)
        db = {
            'code': '200',
            'msg': '成功',
            'fill': 'E7:P9',
            'data': work_data.tolist()
        }
        print(db)
        with open(os.path.join(self.fileDir,'work1.txt'), "w") as fp:
            fp.write(json.dumps(db,ensure_ascii=False, indent=4))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
