# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import sys
from vegatables import Ui_MainWindow
from jiamubiao import *


class MainWindow(QMainWindow, Ui_MainWindow,Programme_jiamu):

	startdatelist = [None,None,None]
	enddatelist = [None,None,None]
	strtext = None
	def __init__(self, parent=None):
		super(MainWindow, self).__init__(parent)
		self.setupUi(self)

		self.starttime.dateChanged['QDate'].connect(self.startdate)
		self.endtime.dateChanged['QDate'].connect(self.enddate)
		self.filename.textEdited['QString'].connect(self.setfilename)
		self.pushButton.clicked.connect(self.createxcel)
		self.pushButton_2.clicked.connect(self.close)

	def startdate(self,QDate):
		datetuple = QDate.getDate()
		self.startdatelist[0] = int(datetuple[0])
		self.startdatelist[1] = int(datetuple[1])
		self.startdatelist[2] = int(datetuple[2])

	def enddate(self,QDate):
		datetuple = QDate.getDate()
		self.enddatelist[0] = int(datetuple[0])
		self.enddatelist[1] = int(datetuple[1])
		self.enddatelist[2] = int(datetuple[2])
	def setfilename(self,strtext):
		self.strtext = strtext + ".xls"
	def createxcel(self):
		if self.strtext is not None:
			Programme_jiamu.get_html(self,self.startdatelist,self.enddatelist,self.strtext)
			QtWidgets.QMessageBox.information(self.pushButton,"标题","Excel生成成功")
		else:
			QtWidgets.QMessageBox.information(self.pushButton,"标题","请输入文件名")





if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)
	myWindow = MainWindow()
	myWindow.show()
	sys.exit(app.exec_())	   
