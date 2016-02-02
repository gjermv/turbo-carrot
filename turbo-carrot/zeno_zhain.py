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


class ZenoZhain(QtGui.QWidget):
    
    def __init__(self):
        super(ZenoZhain, self).__init__()
        self.initUI()
    
    def initUI(self):
        self.fileDir = None
        self.fileName = None
        
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
        dirname = QtGui.QFileDialog.getOpenFileName(self, 'Open file','C:\\python\\zenodata')
        self.le_indata.setText(dirname)
        self.fileName = dirname
    
    def set_outdir(self):
        fname = QtGui.QFileDialog.getExistingDirectory(self,'Export directory','C:\\python\\zenodata')
        self.le_outdir.setText(fname)
        self.fileDir = fname

    def create_report(self):
        print(self.fileDir,self.fileName)
        
        if self.fileDir != None:
            print("Create Report")
            self.createSpreadsheet()
            
            
    def createSpreadsheet(self):
        name = self.le_outname.text()
        filename = self.fileDir +'\\'+ name +'\\Zeno_'+ name+ '.xlsx'
        
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        
        endline = self.addDataToSpreadSheet(worksheet)
        self.initWorkbook(worksheet,endline)
        worksheet.print_area(0, 0, endline, 18)
        workbook.close()
        
        QtGui.QMessageBox.information(self, 'Message', 'Excel report successfully created. {}'.format(filename), buttons=QtGui.QMessageBox.Ok)


    def initWorkbook(self,ws,endline):
        
        #Line 1
        ws.merge_range('A1:R1',"Rail Delivery Lineside Equipment Protection Form")
        #Line 2
        ws.merge_range('A2:B2','Job Name:')
        ws.merge_range('G2:H2','PT Number:')
        ws.merge_range('M2:N2','Recipient Representative:')
        ws.merge_range('C2:F2',"")
        ws.merge_range('I2:L2',"")
        ws.merge_range('O2:R2',"")
        
        #Line 3
        ws.merge_range('A3:B3','Recipient:')
        ws.merge_range('G3:H3','Recipient Number:')
        ws.merge_range('M3:N3','Delivery Date:')
        ws.merge_range('C3:F3',"")
        ws.merge_range('I3:L3',"")
        ws.merge_range('O3:R3',"")
        
        #Line 4
        ws.merge_range('A4:B4','Start Mileage :')
        ws.merge_range('G4:H4','Site Access:')
        ws.merge_range('M4:N4','Site Egress:')
        ws.merge_range('C4:F4',"")
        ws.merge_range('I4:L4',"")
        ws.merge_range('O4:R4',"")

        #Line 5
        ws.merge_range('A5:B5','End Mileage:')
        ws.merge_range('G5:H5','No of Rails & Type:')
        ws.merge_range('M5:N5','3rd Rail:')
        ws.merge_range('C5:F5',"")
        ws.merge_range('I5:L5',"")
        ws.merge_range('O5:R5',"")        

        #Line 6
        ws.merge_range('A6:B6','Delivery Line:')
        ws.merge_range('G6:H6','Chute Direction :')
        ws.merge_range('M6:N6','No of Sites:')
        ws.merge_range('C6:F6',"")
        ws.merge_range('I6:L6',"")
        ws.merge_range('O6:R6',"")
        
        #Line 7
        ws.merge_range('A7:B7','Gradient :')
        ws.merge_range('G7:H7','Curvature:')
        ws.merge_range('M7:N7','Cant:')
        ws.merge_range('C7:F7',"Ascent/Descent/ Both")
        ws.merge_range('I7:L7',"")
        ws.merge_range('O7:R7',"Left / Right")
        
        #Line 8
        ws.merge_range('A8:R9',"Access and Egress Hazards: Trips, slips, and falls")
        
        #Line 10 
        header = ['No','Lat','Lon','Lgth','Chain','Type',]
        for i,item in enumerate(header):
            ws.write(9,i,item)
        
        ws.merge_range('G10:P10',"Track Furniture / Preparation Required / Hazards")
        ws.merge_range('Q10:R10',"Movement")
        ws.write('S10',"Link")
        
        #From endline
        row = endline
        ws.merge_range(row,0,row,6,'Can the delivery to be undertaken in reverse?')
        ws.merge_range(row,7,row,9,'')
        ws.merge_range(row,10,row,15,'Has the rail recipient confirmed the worksite is adequate length?')
        ws.merge_range(row,16,row,18,'')
        
        #From endline +1
        row = endline+1
        ws.merge_range(row,0,row,2,'Site Visit Completed By:')
        ws.merge_range(row,3,row,9,'')
        
        ws.merge_range(row,10,row,11,'Job Title:')
        ws.merge_range(row,12,row,13,'')
        
        ws.merge_range(row,14,row,15,'Date:')
        ws.merge_range(row,16,row,18,'')
        
        #From endline +2
        row = endline+2
        ws.merge_range(row,0,row+1,18,'As the rail recipient, I agree that all preparatory work including equipment protection will be completed before the delivery,trained resources and rail cutting equipment will be available andthe line and adjacent lines within 4m shall be blocked and have no OTM/OTP/Engineering train present.')
        
        #From endline +4
        row = endline+4
        ws.merge_range(row,0,row,2,'Signature:')
        ws.merge_range(row,3,row,9,'')
        
        ws.merge_range(row,10,row,11,'Job Title:')
        ws.merge_range(row,12,row,13,'')
        
        ws.merge_range(row,14,row,15,'Date:')
        ws.merge_range(row,16,row,18,'')
        
        return ws
    
    def addDataToSpreadSheet(self,worksheet):
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
                ws.write_number(k,0,int(data[0]))
                lat = float(data[3])
                lon = float(data[2])
                ws.write_number(k,1,lat)
                ws.write_number(k,2,lon)
                
                if i == 0:
                    ws.write(k,3,0)
                    ws.write(k,4,0)
                else:
                    #ws.write_number(k,3,self.haversine(lon,lat,lon_tmp,lat_tmp))
                    ws.write_number(k,3,self.distance_on_unit_sphere(lat,lon,lat_tmp,lon_tmp))
                    ws.write_number(k,3,self.haversine(lon,lat,lon_tmp,lat_tmp))
                    ws.write_formula(k,4,'='+xl_rowcol_to_cell(k-1,4)+'+'+xl_rowcol_to_cell(k,3))
                
                ws.write(k,5,data[7])
                ws.write(k,6,data[8])
                ws.write(k,16,'----')
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
        r = 6371000 # Radius of earth in meters. Use 3956 for miles
        return c * r

    def distance_on_unit_sphere(self,lat1, long1, lat2, long2):
     
        # Convert latitude and longitude to 
        # spherical coordinates in radians.
        degrees_to_radians = math.pi/180.0
             
        # phi = 90 - latitude
        phi1 = (90.0 - lat1)*degrees_to_radians
        phi2 = (90.0 - lat2)*degrees_to_radians
             
        # theta = longitude
        theta1 = long1*degrees_to_radians
        theta2 = long2*degrees_to_radians
             
        # Compute spherical distance from spherical coordinates.
             
        # For two locations in spherical coordinates 
        # (1, theta, phi) and (1, theta', phi')
        # cosine( arc length ) = 
        #    sin phi sin phi' cos(theta-theta') + cos phi cos phi'
        # distance = rho * arc length
         
        cos = (math.sin(phi1)*math.sin(phi2)*math.cos(theta1 - theta2) + 
               math.cos(phi1)*math.cos(phi2))
        arc = math.acos( cos )
     
        # Remember to multiply arc by the radius of the earth 
        # in your favorite set of units to get length.
        return arc*self.earthRadius(lat1)

    def earthRadius(self,lat):              
        #Rt = radius of earth at latitude t                 
        f = math.radians(lat)
        a = 6378137
        b = 6356752.31420
        
        Rt = math.sqrt(( (a**2* math.cos(f))**2 + (b**2*math.sin(f))**2 ) / ( (a*math.cos(f))**2 + (b*math.sin(f))**2 ))
        
        print('RRT',Rt)
        return Rt         
                 


def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = ZenoZhain()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()





