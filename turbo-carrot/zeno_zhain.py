# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 16:42:41 2015

@author: gjermund.vingerhagen
"""

import sys
from PyQt4 import QtGui,Qt,QtCore
from datetime import datetime as dt
import re



class ZenoZhain(QtGui.QWidget):
    
    def __init__(self):
        super(ZenoZhain, self).__init__()
        self.initUI()
    
    def initUI(self):

        

        # Labels
        self.la_indata = QtGui.QLabel('In data',self)
        self.la_indata.move(15,15)
        
        self.la_outdir = QtGui.QLabel('Directory',self)
        self.la_outdir.move(15,40)

        self.la_outname = QtGui.QLabel('Name',self)
        self.la_outname.move(15,65)
        

        
        #Line edit textboxes
        self.le_indata = QtGui.QLineEdit(self)
        self.le_indata.setFixedWidth(200)
        self.le_indata.move(100,15)
        
        self.le_outdir = QtGui.QLineEdit(self)
        self.le_outdir.setFixedWidth(200)
        self.le_outdir.move(100,40)
        
        self.le_outname = QtGui.QLineEdit(self)
        self.le_outname.setFixedWidth(200)
        self.le_outname.move(100,65)
        
        #Push buttons
        self.btn_indata = QtGui.QPushButton('...', self)
        self.btn_indata.move(305, 14)
        self.btn_indata.setFixedSize(30,20)
        #self.btn1.setStyleSheet("background-color: lightgreen")  
        #self.btn1.clicked.connect(self.btn1Clicked)

        self.btn_outdir = QtGui.QPushButton('...', self)
        self.btn_outdir.move(305, 39)
        self.btn_outdir.setFixedSize(30,20)
        
        self.btn_report = QtGui.QPushButton('Create Report', self)
        self.btn_report.move(15, 90)
        
        #Status bar

        
        #Main window
        self.setGeometry(500,300,345,150)
        self.setWindowTitle('Zeno Zhain')
        self.show()
        
        #Connections
        self.btn_indata.clicked.connect(self.findfile)
        self.btn_outdir.clicked.connect(self.set_outdir)
        self.btn_report.clicked.connect(self.create_report)
        
        
    def findfile(self):
        
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Open file', 
                '/home')
        

        self.le_indata.setText(fname)
    
    def set_outdir(self):
        fname = QtGui.QFileDialog.getExistingDirectory(self,'Choose directory')
        self.le_outdir.setText(fname)

    def create_report(self):
        pass
        
 

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = ZenoZhain()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

