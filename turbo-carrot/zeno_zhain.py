# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 16:42:41 2015

@author: gjermund.vingerhagen
"""

import sys
from PyQt4 import QtGui,Qt,QtCore
from datetime import datetime as dt
import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import zipfile
import glob
import math
import subprocess


class ZenoZhain(QtGui.QWidget):
    
    def __init__(self):
        super(ZenoZhain, self).__init__()
        self.initUI()
        

    
    def initUI(self):
        self.fileDir = None
        self.fileName = None
        
        #Menu bar
        myQMenuBar = QtGui.QMenuBar(self)
        
        exitAction = QtGui.QAction('Exit', self)        
        exitAction.triggered.connect(QtGui.qApp.quit)
        
        helpAction = QtGui.QAction('Help', self)     
        helpAction.triggered.connect(self.showhelp)
        
        aboutAction = QtGui.QAction('About', self)
        aboutAction.triggered.connect(self.showAbout)
        
        myQMenuBar.addAction(exitAction)
        myQMenuBar.addAction(helpAction)
        myQMenuBar.addAction(aboutAction)
        

        
        # Labels
        self.la_indata = QtGui.QLabel('Choose Zeno file',self)
        self.la_indata.move(15,25)
        
        self.la_outdir = QtGui.QLabel('Directory',self)
        self.la_outdir.move(15,50)

        self.la_outname = QtGui.QLabel('Name',self)
        self.la_outname.move(15,75)
        
        #Line edit textboxes
        self.le_indata = QtGui.QLineEdit(self)
        self.le_indata.setFixedWidth(200)
        self.le_indata.move(100,25)
        
        self.le_outdir = QtGui.QLineEdit(self)
        self.le_outdir.setFixedWidth(200)
        self.le_outdir.move(100,50)
        
        self.le_outname = QtGui.QLineEdit(self)
        self.le_outname.setFixedWidth(200)
        self.le_outname.move(100,75)
        
        #Push buttons
        self.btn_indata = QtGui.QPushButton('...', self)
        self.btn_indata.move(305, 24)
        self.btn_indata.setFixedSize(30,20)

        self.btn_outdir = QtGui.QPushButton('...', self)
        self.btn_outdir.move(305, 49)
        self.btn_outdir.setFixedSize(30,20)
        
        self.btn_report = QtGui.QPushButton('Create Report', self)
        self.btn_report.move(15, 100)
        
        
        #Main window
        self.setGeometry(500,300,345,150)
        self.setWindowTitle('Zeno Zhain')
        self.show()
        
        #Connections
        self.btn_indata.clicked.connect(self.findfile)
        self.btn_outdir.clicked.connect(self.set_outdir)
        self.btn_report.clicked.connect(self.create_report)
        
        
    def findfile(self):
        filename = QtGui.QFileDialog.getOpenFileName(self, 'Open file','C:\\python\\zenodata')
        
        if self.checkZenoFileName(filename):
            self.le_indata.setText(filename)
            self.fileName = filename
        else:
             QtGui.QMessageBox.information(self, 'Warning','The file you selected is not recognized', buttons=QtGui.QMessageBox.Ok)
    
    def checkZenoFileName(self,filename):
        print(filename)
        if '.txt' in filename and '.zip' in filename:
            return True
        else:
            return False
    
    def set_outdir(self):
        fname = QtGui.QFileDialog.getExistingDirectory(self,'Export directory','C:\\python\\zenodata')
        self.le_outdir.setText(fname)
        self.fileDir = fname

    def create_report(self):
        print(self.fileDir,self.fileName)
        
        if self.fileDir == None:
            QtGui.QMessageBox.information(self, 'Warning','No new directory is specified', buttons=QtGui.QMessageBox.Ok)
            return None
        elif self.fileName == None:
            QtGui.QMessageBox.information(self, 'Warning','No file is choosed for processing', buttons=QtGui.QMessageBox.Ok)
            return None
        elif self.le_outname.text() == '':
            QtGui.QMessageBox.information(self, 'Warning','No name is specified', buttons=QtGui.QMessageBox.Ok)
            return None            
        if self.fileDir != None:
            print("Create Report")
            self.createSpreadsheet()
            
            
    def createSpreadsheet(self):
        name = self.le_outname.text()
        newdir = self.fileDir +'\\'+ name
        filename = self.fileDir +'\\'+ name +'\\Zeno_'+ name+ '.xlsx'
        
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        
        endline = self.addDataToSpreadSheet(worksheet,workbook)
        self.initWorkbook(worksheet,endline,workbook)
        worksheet.print_area(0, 0, endline, 18)
        workbook.close()
        
        answer = QtGui.QMessageBox.question(self, 'Message', 'Excel report successfully created.\n {}\n\n Do you want to open the folder to view the file?'.format(filename), buttons=QtGui.QMessageBox.Yes, defaultButton=QtGui.QMessageBox.No)
        
        if answer == QtGui.QMessageBox.Yes:
            subprocess.Popen('explorer "{}"'.format(newdir))


    def initWorkbook(self,ws,endline,workbook):
        formatHeader = workbook.add_format({'bold': True, 'font_size': 15,'align':'center'})
        formatStandard = workbook.add_format({'align':'left','border':1,})
        formatStandardPlus = workbook.add_format({'align':'left','border':1,'bold': True})
        formatAlignTop = workbook.add_format({'align':'left','valign':'top','text_wrap':True,'border':1})
        
        #Line 1
        ws.merge_range('A1:R1',"Rail Delivery Lineside Equipment Protection Form",formatHeader)
        #Line 2
        ws.merge_range('A2:B2','Job Name:',formatStandard)
        ws.merge_range('G2:H2','PT Number:',formatStandard)
        ws.merge_range('M2:N2','Recipient Repr:',formatStandard)
        ws.merge_range('C2:F2',"",formatStandard)
        ws.merge_range('I2:L2',"",formatStandard)
        ws.merge_range('O2:R2',"",formatStandard)
        
        #Line 3
        ws.merge_range('A3:B3','Recipient:',formatStandard)
        ws.merge_range('G3:H3','Recipient Number:',formatStandard)
        ws.merge_range('M3:N3','Delivery Date:',formatStandard)
        ws.merge_range('C3:F3',"",formatStandard)
        ws.merge_range('I3:L3',"",formatStandard)
        ws.merge_range('O3:R3',"",formatStandard)
        
        #Line 4
        ws.merge_range('A4:B4','Start Mileage :',formatStandard)
        ws.merge_range('G4:H4','Site Access:',formatStandard)
        ws.merge_range('M4:N4','Site Egress:',formatStandard)
        ws.merge_range('C4:F4',"",formatStandard)
        ws.merge_range('I4:L4',"",formatStandard)
        ws.merge_range('O4:R4',"",formatStandard)

        #Line 5
        ws.merge_range('A5:B5','End Mileage:',formatStandard)
        ws.merge_range('G5:H5','No of Rails & Type:',formatStandard)
        ws.merge_range('M5:N5','3rd Rail:',formatStandard)
        ws.merge_range('C5:F5',"",formatStandard)
        ws.merge_range('I5:L5',"",formatStandard)
        ws.merge_range('O5:R5',"",formatStandard)        

        #Line 6
        ws.merge_range('A6:B6','Delivery Line:',formatStandard)
        ws.merge_range('G6:H6','Chute Direction :',formatStandard)
        ws.merge_range('M6:N6','No of Sites:',formatStandard)
        ws.merge_range('C6:F6',"",formatStandard)
        ws.merge_range('I6:L6',"",formatStandard)
        ws.merge_range('O6:R6',"",formatStandard)
        
        #Line 7
        ws.merge_range('A7:B7','Gradient :',formatStandard)
        ws.merge_range('G7:H7','Curvature:',formatStandard)
        ws.merge_range('M7:N7','Cant:',formatStandard)
        ws.merge_range('C7:F7',"Ascent/Descent/ Both")
        ws.merge_range('I7:L7',"",formatStandard)
        ws.merge_range('O7:R7',"Left / Right")
        
        #Line 8
        ws.merge_range('A8:R9',"Access and Egress Hazards: Trips, slips, and falls",formatAlignTop)
        
        #Line 10 
        header = ['No','Lat','Lon','Lgth','Chain','Type',]
        for i,item in enumerate(header):
            ws.write(9,i,item,formatStandard)
        
        ws.merge_range('G10:P10',"Track Furniture / Preparation Required / Hazards",formatStandard)
        ws.merge_range('Q10:R10',"Movement",formatStandard)
        ws.write('S10',"Link",formatStandard)
        
        #From endline
        row = endline
        ws.merge_range(row,0,row,6,'Can the delivery to be undertaken in reverse?',formatStandard)
        ws.merge_range(row,7,row,8,'',formatStandard)
        ws.merge_range(row,9,row,15,'Has the rail recipient confirmed the worksite is adequate length?',formatStandard)
        ws.merge_range(row,16,row,17,'',formatStandard)
        
        #From endline +1
        row = endline+1
        ws.merge_range(row,0,row,2,'Site Visit Completed By:',formatStandardPlus)
        ws.merge_range(row,3,row,9,'',formatStandard)
        
        ws.merge_range(row,10,row,11,'Job Title:',formatStandardPlus)
        ws.merge_range(row,12,row,13,'',formatStandard)
        
        ws.merge_range(row,14,row,15,'Date:',formatStandardPlus)
        ws.merge_range(row,16,row,17,'',formatStandard)
        
        #From endline +2
        row = endline+2
        ws.merge_range(row,0,row+1,17,'As the rail recipient, I agree that all preparatory work including equipment protection will be completed before the delivery,trained resources and rail cutting equipment will be available and the line and adjacent lines within 4m shall be blocked and have no OTM/OTP/Engineering train present.',formatAlignTop)
        
        #From endline +4
        row = endline+4
        ws.merge_range(row,0,row,2,'Signature:',formatStandardPlus)
        ws.merge_range(row,3,row,9,'',formatStandard)
        
        ws.merge_range(row,10,row,11,'Job Title:',formatStandardPlus)
        ws.merge_range(row,12,row,13,'',formatStandard)
        
        ws.merge_range(row,14,row,15,'Date:',formatStandardPlus)
        ws.merge_range(row,16,row,17,'',formatStandard)
        
        return ws
    
    def addDataToSpreadSheet(self,worksheet,workbook):
        formatStandard = workbook.add_format({'align':'left','border':1})
        
        ws = worksheet
        k = 0
        
        name = self.le_outname.text()
        exportDir = self.fileDir +'\\'+ name 
        
        ZipFile = zipfile.ZipFile(self.fileName)
        ZipFile.extractall(exportDir)
        ZipFile.close()
        
        subdir = glob.glob(exportDir+'\\*\\')[0]
        
        myFile = glob.glob(exportDir+'\\*\\*.txt')
        print(exportDir,myFile)          
        
        for file in myFile:
            print(file)
            myf = open(file,'r',encoding = 'utf-16') 
            lat_tmp = 0
            lon_tmp = 0
            for i,line in enumerate(myf.readlines()[1:]):
                k = i+10
                data = line.replace('\n','').split(';')
                ws.write_number(k,0,int(data[0]),formatStandard)

                lat = round(float(data[3]),9)
                lon = round(float(data[2]),9)
                ws.write_number(k,1,lat,formatStandard)
                ws.write_number(k,2,lon,formatStandard)
                
                if i == 0:
                    ws.write(k,3,0)
                    ws.write(k,4,0)
                else:
                    ws.write_number(k,3,self.haversine(lon,lat,lon_tmp,lat_tmp),formatStandard)
                    ws.write_formula(k,4,'='+xl_rowcol_to_cell(k-1,4)+'+'+xl_rowcol_to_cell(k,3),formatStandard)
                
                ws.write(k,5,data[7],formatStandard)
                ws.merge_range(k,6,k,15,data[8],formatStandard)
                ws.merge_range(k,16,k,17,'----',formatStandard)
                
                if len(data[5])> 1:
                    ws.write_url(k,18,subdir+'Images\\'+data[5])
                
                lat_tmp,lon_tmp = lat,lon
        return k+2
    
    def haversine(self,lon1, lat1, lon2, lat2):
        """
        Calculate the great circle distance between two points 
        on the earth (specified in decimal degrees)
        """
        # convert decimal degrees to radians 
        lon1, lat1, lon2, lat2 = map(math.radians, [lon1, lat1, lon2, lat2])
    
        # haversine formula 
        dlon = lon2 - lon1 
        dlat = lat2 - lat1 
        a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
        c = 2 * math.asin(math.sqrt(a)) 
        r = 6378100  #Radius of earth in meters.
        return c * r

    def showhelp(self):
        QtGui.QMessageBox.information(self, "Help",
                    """Choose Zeno file -> The exported ascii file from the Zeno controller.\n
Output dir -> The new directory where you want to store the data, ie. Desktop\n
File name -> The name of the folder and exported Excel file.""")

    def showAbout(self):
        QtGui.QMessageBox.information(self, "About",
                    "Version: \napp 0.0.02")

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = ZenoZhain()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()





