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
        self.initUI()

    def initUI(self):
        self.resize(400, 250)
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

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(title, 1, 0)
        grid.addWidget(self.titleEdit, 1, 1)
        grid.addWidget(titleBtn, 1, 2)

        grid.addWidget(work1Btn, 2, 0)
        grid.addWidget(work2Btn, 2, 1)
        grid.addWidget(work3Btn, 2, 2)
        grid.addWidget(work4Btn, 3, 0)
        grid.addWidget(work5Btn, 3, 1)
        grid.addWidget(work6Btn, 3, 2)
        grid.addWidget(work7Btn, 4, 0)

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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
                'num_sheet': len(excel_list),
                'fill': 'D7:M10',
                'data': work_data.tolist()
                # 'data': row_dict
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
                'num_sheet': len(excel_list),
                'fill': 'D6:L8',
                'data': work_data.tolist()
                # 'data': row_dict
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
                'num_sheet': len(excel_list),
                'fill': 'D7:N10',
                'data': work_data.tolist()
                # 'data': row_dict
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
            for file_excel in excel_list:
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
            db = {
                'code': '200',
                'msg': '成功',
                'num_sheet': len(excel_list),
                'fill': 'C7:J10',
                'data': work_data.tolist()
                # 'data': row_dict
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
