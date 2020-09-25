# !/usr/bin/python3
# -*- coding: utf-8 -*-

"""
@Author: 闫继龙
@Version: ??
@License: Apache Licence
@CreateTime: 2019/8/26 12:35
@Describe：

"""

import sys, os

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']

from PyQt5.QtWidgets import (QWidget, QGridLayout, QLabel, QLineEdit, QMessageBox,
                             QTextEdit, QPushButton, QApplication, QDesktopWidget)
import xlrd  # 读取excel
import numpy as np
import json


def all_path(dirname):
    result = []
    for maindir, subdir, file_name_list in os.walk(dirname):
        for filename in file_name_list:

            if is_number(filename.split('.')[0]):
                apath = os.path.join(maindir, filename)
                result.append(apath)

    # print(result)
    # 列表排序
    result.sort(key=lambda x: int(x.split('\\')[-1].split('.')[0]))
    # print(len(result))
    return result  # 返回文件列表


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


def showDialog(self, workname):
    QMessageBox.information(self, "SUCCESS", workname + "汇总完毕",
                            QMessageBox.Yes)


def showErrorDialog(self, workname, errmsg):
    QMessageBox.information(self, "ERROR", workname + ':' + errmsg,
                            QMessageBox.Yes)


class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.fileDir = 'D:\home\partydb'
        self.free_fileDir = r'D:\home\partydb\表1'  # 自由模式下的文件路径
        self.free_row = []  # 自由模式下的行
        self.free_col = []  # 自由模式下的列
        self.initUI()

    def initUI(self):
        self.resize(600, 300)
        self.setWindowTitle("party")

        title = QLabel('fileDir')
        self.titleEdit = QLineEdit(self.fileDir)
        titleBtn = QPushButton('submit', self)
        titleBtn.clicked.connect(self.buttonClicked)

        # 处理btn
        work1Btn = QPushButton('表1', self)
        work1Btn.clicked.connect(self.work1)
        work2Btn = QPushButton('表2', self)
        work2Btn.clicked.connect(self.work2)
        work3Btn = QPushButton('表3', self)
        work3Btn.clicked.connect(self.work3)
        work4Btn = QPushButton('表4', self)
        work4Btn.clicked.connect(self.work4)
        work5Btn = QPushButton('表5', self)
        work5Btn.clicked.connect(self.work5)
        work6Btn = QPushButton('表6', self)
        work6Btn.clicked.connect(self.work6)
        work7Btn = QPushButton('表7', self)
        work7Btn.clicked.connect(self.work7)

        # 自由模式
        free_title = QPushButton('自由模式')
        self.free_titleEdit = QLineEdit(self.free_fileDir)
        self.free_titleRow = QLineEdit('从1开始，输入开始和结束的行号，以空格隔开')
        self.free_titleCol = QLineEdit('从1开始，输入开始和结束的列号，以空格隔开')
        self.free_errorCol = QLineEdit('输入错误的列号，以空格隔开，如果没有则输入空')

        free_titleBtn = QPushButton('submit', self)
        free_titleBtn.clicked.connect(self.free_buttonClicked)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(title, 1, 0)
        grid.addWidget(self.titleEdit, 1, 1, 1, 4)
        grid.addWidget(titleBtn, 1, 6)

        grid.addWidget(work1Btn, 2, 0)
        grid.addWidget(work2Btn, 2, 1)
        grid.addWidget(work3Btn, 2, 2)
        grid.addWidget(work4Btn, 2, 3)
        grid.addWidget(work5Btn, 2, 4)
        grid.addWidget(work6Btn, 2, 5)
        grid.addWidget(work7Btn, 2, 6)

        # 自由模式
        grid.addWidget(free_title, 3, 1, 1, 5)
        grid.addWidget(self.free_titleEdit, 4, 0, 1, 3)
        grid.addWidget(self.free_titleRow, 4, 4, 1, 3)
        grid.addWidget(self.free_titleCol, 6, 0, 1, 3)
        grid.addWidget(self.free_errorCol, 6, 4, 1, 3)
        grid.addWidget(free_titleBtn, 7, 1, 1, 5)

        self.setLayout(grid)
        self.center()
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

    def free_buttonClicked(self):
        sender = self.sender()
        rootdir = self.free_titleEdit.text()
        free_row = list(map(int, self.free_titleRow.text().split(' ')))
        free_col = list(map(int, self.free_titleCol.text().split(' ')))
        free_row = [int(i) - 1 for i in range(free_row[0], free_row[1]+1)]
        free_col = [int(i) - 1 for i in range(free_col[0], free_col[1]+1)]

        if self.free_errorCol.text() == '':
            col_err_list = []
        else:
            col_err_list = [int(i) - 1 for i in self.free_errorCol.text().split(' ')]
        excel_list = all_path(rootdir)

        try:
            # 所有数据
            work_data = np.zeros((len(free_row), len(free_col)), dtype='int64')
            xml_list = []
            minus_num = []  # 负数表

            for file_excel in excel_list:
                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)

                sheet_data = []  # 表数据
                # 按行遍历一张表
                for row_num in free_row:
                    row = sheet.row_values(row_num)
                    row_data = []  # 行数据
                    for c in free_col:
                        df = row[c]

                        if df == "":
                            df = 0
                        elif isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif c in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'{rootdir[-2:]}:{file_excel},{row_num + 1}行,{c + 1}列', f'存在异常数据{df}')
                            return
                        if int(df) < 0:
                            minus_num.append(file_excel.split('\\')[-1])

                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)

                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'work'+ rootdir[-1]+f':{file_excel}', '存在小于0的值')
                    return
                    # print(sheet_data)
                work_data += np.array(sheet_data)
                if np.any(work_data < 0):
                    showErrorDialog(self, f'work'+ rootdir[-1]+f':{file_excel}', '存在小于0的值')
                    return
            # print(work_data)

            db = {
                'code': '000',
                'msg': '查看信息',
                'describe': 'xml_sheet和minus_seet，可能会对计算结果带来影响，请手动计算err_sheet表',
                'xml_sheet(xml文件列表)': xml_list,
                'minus_seet(存在负数的文件)': minus_num,
                'num_sheet(程序正确处理的文件)': len(excel_list),
                'fill': chr(free_col[0] + 65) + str(free_row[0] + 1) + ':' + chr(
                    free_col[-1] + 65) + str(free_row[-1] + 1),
                'data': work_data.tolist()
            }
            # print(db)
            with open(os.path.join(self.fileDir, 'work' + rootdir[-1] + '.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, str(rootdir[-2:]))
        except Exception as e:
            print(e)
            showErrorDialog(self, 'work' + str(rootdir[-1]), e.__str__())

    def work1(self):
        try:
            sender = self.sender()

            rootdir = os.path.join(self.fileDir, '表1')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((3, 12), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件3—党组织情况统计表（表1）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
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
                    col_err_list = []
                    for c in range(len(col_list)):
                        df = row[col_list[c]]

                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 4 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'work1:{file_excel}', '存在小于0的值')
                    return
                    # print(sheet_data)
                work_data += sheet_data
                # print('88888888888888' * 8)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'E7:P9',
                    'data': work_data.tolist()
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'E7:P9',
                    'data': work_data.tolist()
                }
            # print(db)
            with open(os.path.join(self.fileDir, 'work1.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work1')
        except Exception as e:
            showErrorDialog(self, 'work1', e.__str__())

    def work2(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表2')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((10, 19), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件4—教师党员结构统计表（表2）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据
                # data_sheet = np.zeros((3, 12))
                # 按行遍历一张表
                for row_num in range(7, 17):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row

                    row_data = []  # 行数据
                    col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]
                    col_err_list = [3, 5, 10, 14, 21]
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 3 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表2:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data += sheet_data
                # print('88888888888888' * 8)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'D8:V17',
                    'data': work_data.tolist()
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'D8:V17',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work2.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work2')
        except Exception as e:
            showErrorDialog(self, 'work2', e.__str__())

    def work3(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表3')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((4, 10), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件5—高层次人才党员统计表（表3）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据

                # 按行遍历一张表
                for row_num in range(6, 10):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row
                    # print(row[3:13])
                    row_data = []  # 行数据
                    col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
                    col_err_list = []
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 3 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表3:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data = work_data + sheet_data

                # print('88888888888888' * 8)
            # print(work_data)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'D7:M10',
                    'data': work_data.tolist()
                    # 'data': row_dict
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'D7:M10',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work3.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work3')
        except Exception as e:
            showErrorDialog(self, 'work3', e.__str__())

    def work4(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表4')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((3, 9), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件6—“双带头人”党支部书记配备情况统计表（表4）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据

                # 按行遍历一张表
                for row_num in range(5, 8):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row
                    # print(row[3:13])
                    row_data = []  # 行数据
                    col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11]
                    col_err_list = []
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 3 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表4:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data = work_data + sheet_data

                # print('88888888888888' * 8)
            # print(work_data)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'D6:L8',
                    'data': work_data.tolist()
                    # 'data': row_dict
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'D6:L8',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work4.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work4')
        except Exception as e:
            showErrorDialog(self, 'work4', e.__str__())

    def work5(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表5')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((16, 10), dtype='int16')
            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件7—学生党员结构和党组织统计表（表5）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据
                # data_sheet = np.zeros((3, 12))
                # 按行遍历一张表
                for row_num in range(7, 23):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row

                    row_data = []  # 行数据
                    col_list = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
                    col_err_list = [4, 9, 13]
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 4 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表5:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data += sheet_data
                # print('88888888888888' * 8)
            # print(work_data)
            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'E8:N23',
                    'data': work_data.tolist()
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'E8:N23',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work5.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work5')
        except Exception as e:
            showErrorDialog(self, 'work5', e.__str__())

    def work6(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表6')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((4, 11), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件8—失联党员情况汇总表（表6）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据

                # 按行遍历一张表
                for row_num in range(6, 10):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row
                    # print(row[3:13])
                    row_data = []  # 行数据
                    col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
                    col_err_list = []
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 3 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表6:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data = work_data + sheet_data
                # print('88888888888888' * 8)
            # print(work_data)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'D7:N10',
                    'data': work_data.tolist()
                    # 'data': row_dict
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'D7:N10',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work6.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work6')
        except Exception as e:
            showErrorDialog(self, 'work6', e.__str__())

    def work7(self):
        try:
            sender = self.sender()
            rootdir = os.path.join(self.fileDir, '表7')
            excel_list = all_path(rootdir)
            # print(excel_list)
            row_dict = {}

            # 所有数据
            work_data = np.zeros((4, 8), dtype='int16')

            # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
            xml_list = []
            for file_excel in excel_list:

                # 判断文本格式，如果‘xml’是错误文件
                with open(file_excel, 'r', encoding='ISO-8859-1') as f:
                    line = f.readline()
                    # print(line)
                    if 'xml' in line:
                        xml_list.append(file_excel.split('\\')[-1])

                # by_sheet = u'附件9-失联党员组织处置情况汇总表（表7）'
                # 打开数据所在的工作簿，以及选择存有数据的工作表
                book = xlrd.open_workbook(file_excel)
                by_sheet = book.sheets()[0].name
                sheet = book.sheet_by_name(by_sheet)
                n_rows = sheet.nrows  # 行数

                sheet_data = []  # 表数据

                # 按行遍历一张表
                for row_num in range(6, 10):
                    row = sheet.row_values(row_num)
                    row_dict[row_num] = row
                    # print(row[3:13])
                    row_data = []  # 行数据
                    col_list = [2, 3, 4, 5, 6, 7, 8, 9]
                    col_err_list = []
                    for c in range(len(col_list)):
                        df = row[col_list[c]]
                        if isinstance(df, str):
                            df = 0
                        elif isinstance(df, float):
                            df = int(df)
                        elif col_list[c] in col_err_list:
                            df = 0
                        else:
                            showErrorDialog(self, f'表2:{file_excel},{row_num + 1}行,{c + 2 + 1}列', f'存在异常数据{df}')
                            return
                        row_data.append(df)
                    sheet_data.append(row_data)

                sheet_data = np.array(sheet_data)
                if np.any(sheet_data < 0):
                    showErrorDialog(self, f'表7:{file_excel}', '存在小于0的值')
                    return
                # print(sheet_data)
                work_data = work_data + sheet_data
                # print('88888888888888' * 8)
            # print(work_data)

            db = {}
            if len(xml_list) == 0:
                db = {
                    'code': '200',
                    'msg': '成功',
                    'num_sheet': len(excel_list),
                    'fill': 'C7:J10',
                    'data': work_data.tolist()
                    # 'data': row_dict
                }
            else:
                db = {
                    'code': '500',
                    'msg': '失败',
                    'describe': '以下文件是xml文本，可能会对计算结果带来影响，请手动计算err_sheet表',
                    'err_sheet': xml_list,
                    'num_sheet': len(excel_list),
                    'fill': 'C7:J10',
                    'data': work_data.tolist()
                }
            with open(os.path.join(self.fileDir, 'work7.txt'), "w") as fp:
                fp.write(json.dumps(db, ensure_ascii=False, indent=4))
            showDialog(self, 'work7')
        except Exception as e:
            showErrorDialog(self, 'work7', e.__str__())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
