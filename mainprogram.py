# -*- coding:utf-8 -*-
from main import Ui_Form
# from address import Ui_Form
from PyQt5 import QtCore, QtGui, QtWidgets
import sys

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)

    MainWindow = QtWidgets.QMainWindow()

    ui = Ui_Form()

    ui.setupUi(MainWindow)

    MainWindow.show()

sys.exit(app.exec_())
