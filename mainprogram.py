# -*- coding:utf-8 -*-
from addresys import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets, QtSql
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from addData import Ui_dialog
import sys
import pymysql
import xlwt
from PyQt5 import sip

class maininterface(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(maininterface, self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.__addData)
        self.pushButton_2.clicked.connect(self.__search)
        self.pushButton_6.clicked.connect(self.__connDb)
        self.pushButton_7.clicked.connect(self.__closeconn)
        self.pushButton_4.clicked.connect(self.__exportData)
        self.lineEdit_9.setEchoMode(QLineEdit.Password)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # self.tableWidget.cellChanged.connect(self.cellchange)
        self.conn = None

    def __exportData(self):
        fileName2, ok2 = QFileDialog.getSaveFileName(self,
                                                     "文件保存",
                                                     "C:/",
                                                     "Excel Files (*.xls)")
        # print(fileName2, ok2)
        if fileName2 and ok2:
            wbk = xlwt.Workbook()
            sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)
            title = ['部门', '姓名', '职务', '手机', '短号', '办公电话']
            for i in range(0, len(title)):
                sheet.write(0, i, title[i])
            for i in range(len(self.result)):
                for j in range(1, len(self.result[i])):
                    sheet.write(i + 1, j - 1, self.result[i][j])
            wbk.save(fileName2)

    # 列表内添加按钮
    def buttonForRow(self, id, i):
        widget = QWidget()
        # 修改
        updateBtn = QPushButton('修改')
        updateBtn.setStyleSheet(''' text-align : center;
                                              background-color : NavajoWhite;
                                              height : 30px;
                                              border-style: outset;
                                              font : 13px  ''')

        updateBtn.clicked.connect(lambda: self.updateTable(id, i))

        # 删除
        deleteBtn = QPushButton('删除')
        deleteBtn.setStyleSheet(''' text-align : center;
                                        background-color : LightCoral;
                                        height : 30px;
                                        border-style: outset;
                                        font : 13px; ''')
        deleteBtn.clicked.connect(lambda: self.deleteTable(id))
        hLayout = QHBoxLayout()
        hLayout.addWidget(updateBtn)
        hLayout.addWidget(deleteBtn)
        hLayout.setContentsMargins(5, 2, 5, 2)
        widget.setLayout(hLayout)
        return widget

    def cellchange(self, row, col):
        print(self.tableWidget.item(row, col).text())
        self.result[row][col + 1] = self.tableWidget.item(row, col).text()

    def updateTable(self, id, row):
        self.tableWidget.columnCount()
        department = self.tableWidget.item(row, 0).text().strip()
        name = self.tableWidget.item(row, 1).text().strip()
        duty = self.tableWidget.item(row, 2).text().strip()
        phone = self.tableWidget.item(row, 3).text().strip()
        shortnum = self.tableWidget.item(row, 4).text().strip()
        officephone = self.tableWidget.item(row, 5).text().strip()
        sql = 'update addressBook set department="{}",name="{}", duty="{}", phone="{}", shortnum="{}", officephone="{}" where id={}'.format(
            department, name, duty, phone, shortnum, officephone, id)
        try:
            self.cur.execute(sql)
            self.conn.commit()
            self.__search()
            QMessageBox.warning(self, '提示', '修改成功', QMessageBox.Ok)
        except Exception as e:
            self.conn.rollback()
            QMessageBox.warning(self, '警告', str(e.args), QMessageBox.Cancel)

    def deleteTable(self, id):
        print(id)
        rep = QMessageBox.question(self, '提示', '确定删除吗？', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if rep == QMessageBox.Yes:
            try:
                self.cur.execute('delete from addressBook where id ={};'.format(id))
                self.conn.commit()
                self.showAllData()
            except Exception as e:
                self.conn.rollback()
                QMessageBox.warning(self, '警告', str(e.args), QMessageBox.Cancel)

    def __closeconn(self):
        self.conn.close()
        self.pushButton.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.pushButton_4.setEnabled(False)
        self.pushButton_7.setEnabled(False)
        self.pushButton_6.setEnabled(True)
        self.lineEdit_7.setEnabled(True)
        self.lineEdit_8.setEnabled(True)
        self.lineEdit_9.setEnabled(True)
        self.tableWidget.setRowCount(0)
        print('close conn')

    def __search(self):
        department = self.lineEdit.text().strip()
        name = self.lineEdit_2.text().strip()
        duty = self.lineEdit_3.text().strip()
        phone = self.lineEdit_4.text().strip()
        shortnum = self.lineEdit_5.text().strip()
        officephone = self.lineEdit_6.text().strip()
        if not department and not name and not duty and not phone and not shortnum and not officephone:
            self.showAllData()
        else:
            sql = 'select * from addressBook where 1'
            if department:
                sql += ' and department like "%{}%"'.format(department)
            if name:
                sql += ' and name like "%{}%"'.format(name)
            if duty:
                sql += ' and duty like "%{}%"'.format(duty)
            if phone:
                sql += ' and phone like "%{}%"'.format(phone)
            if shortnum:
                sql += ' and shortnum like "%{}%"'.format(shortnum)
            if officephone:
                sql += ' and officephone like "%{}%"'.format(officephone)
            print(sql)
            self.showAllData(sql)

    def closeEvent(self, QCloseEvent):
        rep = QMessageBox.question(self, '提示', '确定退出吗？', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if rep == QMessageBox.Yes:
            if self.conn and self.pushButton_7.isEnabled():
                self.conn.close()
                print('close conn')
            QCloseEvent.accept()
        else:
            QCloseEvent.ignore()

    def __connDb(self):
        ip = self.lineEdit_7.text().strip()
        user = self.lineEdit_8.text().strip()
        pwd = self.lineEdit_9.text().strip()
        if ip and user and pwd:
            try:
                self.conn = pymysql.Connect(host=ip, user=user, password=pwd, database='lws')
                if self.conn:
                    self.cur = self.conn.cursor()
                    self.pushButton.setEnabled(True)
                    self.pushButton_2.setEnabled(True)
                    self.pushButton_3.setEnabled(True)
                    self.pushButton_4.setEnabled(True)
                    self.pushButton_7.setEnabled(True)
                    self.pushButton_6.setEnabled(False)
                    self.lineEdit_7.setEnabled(False)
                    self.lineEdit_8.setEnabled(False)
                    self.lineEdit_9.setEnabled(False)
                    self.showAllData()
            except Exception as e:
                QMessageBox.warning(self, '警告', str(e.args), QMessageBox.Cancel)
        else:
            QMessageBox.warning(self, '警告', '输入有误！', QMessageBox.Cancel)

    def showAllData(self, sql=None):
        if sql == None:
            sql = 'SELECT * from addressBook;'
        self.cur.execute(sql)
        self.result = self.cur.fetchall()
        row = self.cur.rowcount
        vol = len(self.result[0])
        self.tableWidget.setRowCount(row)
        for i in range(row):
            for j in range(1, vol):
                temp_data = self.result[i][j]
                data = QTableWidgetItem(str(temp_data))
                self.tableWidget.setItem(i, j - 1, data)
                if j == vol - 1:
                    self.tableWidget.setCellWidget(i, j, self.buttonForRow(self.result[i][0], i))
        # print(self.result)

    def __addData(self):
        self.adddata_dialog = addDataDialog()
        self.adddata_dialog.sql_signal.connect(self.__insertdata)
        self.adddata_dialog.exec_()

    def __insertdata(self, sql):
        try:
            self.cur.execute(sql)
            self.conn.commit()
            self.adddata_dialog.close()
            self.showAllData()
        except Exception as e:
            self.conn.rollback()
            QMessageBox.warning(self, '警告', str(e.args), QMessageBox.Cancel)


class addDataDialog(Ui_dialog, QDialog):
    sql_signal = pyqtSignal(str)

    def __init__(self):
        super(addDataDialog, self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.__adddata)

    def __adddata(self):
        if self.lineEdit_2.text().strip() and self.lineEdit_4.text().strip():
            department = self.lineEdit.text().strip()
            name = self.lineEdit_2.text().strip()
            duty = self.lineEdit_3.text().strip()
            phone = self.lineEdit_4.text().strip()
            shortnum = self.lineEdit_5.text().strip()
            officephone = self.lineEdit_6.text().strip()
            sql = "INSERT INTO addressBook(department,NAME,duty,phone,shortnum,officephone) VALUES('{}','{}','{}','{}','{}','{}');".format(
                department, name, duty, phone, shortnum, officephone)
            self.sql_signal.emit(sql)
        else:
            QMessageBox.warning(self, '警告', '姓名和手机必填！', QMessageBox.Cancel)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)

    MainWindow = maininterface()

    MainWindow.show()

    sys.exit(app.exec_())
