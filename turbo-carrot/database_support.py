# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 16:42:41 2015

@author: gjermund.vingerhagen
"""

import sys
from PyQt4 import QtGui,Qt
import sqlite3 as lite
from datetime import datetime as dt
import re

class TogglCopy(QtGui.QWidget):
    
    def __init__(self):
        super(TogglCopy, self).__init__()
        self.initUI()
        self.btn1Flag = False
        self.startTime = 0
        
    def initUI(self):
        # Labels
        self.label1 = QtGui.QLabel('Register case',self)
        self.label1.move(20,10)
        
        self.label2 = QtGui.QLabel('Start Time',self)
        self.label2.move(130,50)

        self.label3 = QtGui.QLabel('Duration',self)
        self.label3.move(230,50)
        
        self.label4 = QtGui.QLabel('Name',self)
        self.label4.move(30,100)

        self.label5 = QtGui.QLabel('Company',self)
        self.label5.move(130,100)
            
        self.label6 = QtGui.QLabel('Phone number',self)
        self.label6.move(230,100)
        
        self.label7 = QtGui.QLabel('Products',self)
        self.label7.move(30,155)

        self.label8 = QtGui.QLabel('Equipment no',self)
        self.label8.move(230,155)
        
        self.label9 = QtGui.QLabel('Short description of the problem',self)
        self.label9.move(30,205)
        
        self.label9 = QtGui.QLabel('Solution',self)
        self.label9.move(30,300)
        
        #Line edit textboxes
        self.le_StartTime = QtGui.QLineEdit(self)
        self.le_StartTime.setFixedWidth(90)
        self.le_StartTime.move(130,65)
        
        self.le_Duration = QtGui.QLineEdit(self)
        self.le_Duration.setFixedWidth(90)        
        self.le_Duration.move(230,65)
        
        self.le_Name = QtGui.QLineEdit(self)
        self.le_Name.setFixedWidth(90)
        self.le_Name.move(30,115)
        
        self.le_Company = QtGui.QLineEdit(self)
        self.le_Company.setFixedWidth(90)
        self.le_Company.move(130,115)
        
        self.le_Phone = QtGui.QLineEdit(self)
        self.le_Phone.setFixedWidth(80)
        self.le_Phone.move(230,115)
        
        self.le_Products = QtGui.QLineEdit(self)
        self.le_Products.setFixedWidth(190)
        self.le_Products.move(30,170)

        self.le_EquipmentNo = QtGui.QLineEdit(self)
        self.le_EquipmentNo.setFixedWidth(90)
        self.le_EquipmentNo.move(230,170)
        
        #Textboxes
        self.te_Problem = QtGui.QTextEdit(self)
        self.te_Problem.setFixedSize(290,80)
        self.te_Problem.setTabChangesFocus(True)
        self.te_Problem.move(30, 220)
        
        self.te_Solution = QtGui.QTextEdit(self)
        self.te_Solution.setTabChangesFocus(True)
        self.te_Solution.setFixedSize(290,80)
        self.te_Solution.move(30, 315)

        #Checkboxes
        self.cb_Solved = QtGui.QCheckBox('Solved', self)
        self.cb_Solved.move(30,405)
        
        self.cb_Forwarded = QtGui.QCheckBox('Forwarded', self)
        self.cb_Forwarded.move(130,405)        
        
        self.cb_FollowUp = QtGui.QCheckBox('Follow up', self)
        self.cb_FollowUp.move(230,405)

        #Push buttons
        self.btn1 = QtGui.QPushButton('Start Time', self)
        self.btn1.setFixedSize(90,55)
        self.btn1.move(30, 35)
        self.btn1.setStyleSheet("background-color: lightgreen")  
        self.btn1.clicked.connect(self.btn1Clicked)
        
        self.btn2 = QtGui.QPushButton('Register', self)
        self.btn2.setFixedWidth(60)
        self.btn2.move(30, 430)
        self.btn2.clicked.connect(self.btn2Clicked)
        
        self.btn3 = QtGui.QPushButton('Clear', self)
        self.btn3.move(100, 430)
        self.btn3.clicked.connect(self.btn3Clicked)
        
        self.btn4 = QtGui.QPushButton('?', self)
        self.btn4.setFixedWidth(20)
        self.btn4.move(310, 113)
        #self.btn4.clicked.connect(self.btn3Clicked)
        
        #Status bar
        self.statuslabel = QtGui.QLabel('     Ready',self)
        self.statuslabel.move(30,470)
        self.statuslabel.setFixedSize(290,22)
        self.statuslabel.setStyleSheet("background-color:silver; ")
        
        
        
        #Main window
        self.setGeometry(500,300,345,500)
        self.setWindowTitle('Time registration')
        self.show()
        
        self.completer = QtGui.QCompleter()
        self.le_Company.setCompleter(self.completer)
        
        model = QtGui.QStringListModel()
        self.completer.setModel(model)
        self.get_data(model)
    
    def btn1Clicked(self):
        if not self.btn1Flag:
            self.startTime = dt.now()
            self.le_StartTime.setText(str(self.startTime.strftime("%H:%M")))
            self.btn1.setText("Stop Time")
            self.btn1.setStyleSheet("background-color: red")
            self.btn1Flag = True
            
        else:
            durtime = str((dt.now()-self.startTime)).split('.')[0]
            self.le_Duration.setText(durtime)
            self.btn1.setText("Start Time")
            self.btn1.setStyleSheet("background-color: lightgreen")  
            self.btn1Flag = False
            
            

    def btn2Clicked(self):
        self.con = lite.connect('C:\\python\\database\\test_X.db')
        
        startTime = self.startTime       
        duration = self.le_Duration.text()
        name = self.le_Name.text()
        company = self.le_Company.text()
        phone = self.le_Phone.text()
        products = self.le_Products.text()
        equipmentno = self.le_EquipmentNo.text()
        problem = self.te_Problem.toPlainText()
        solution = self.te_Solution.toPlainText()
        regtime = str(dt.now())
        
        if self.cb_Solved.checkState() == 2:
            solved = 1
        else: 
            solved = 0
        if self.cb_FollowUp.checkState() == 2:
            followup = 1
        else:
            followup = 0
        if self.cb_Forwarded.checkState() == 2:
            forwarded = 1
        else:
            forwarded = 0
        
        exestring = "INSERT INTO Support VALUES(NULL,"
        exestring += "'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')".format(startTime, duration, regtime, name, company, phone, products, equipmentno, problem, solution,solved,followup,forwarded)
        print(exestring)
        with self.con:
            cur = self.con.cursor()        
            cur.execute(exestring)
                
        self.con.commit()
        self.con.close()
        
        self.statuslabel.setText(' -Successfully stored in database.')
        
        self.messagebox =  QtGui.QMessageBox.question(self, 'Message',
            "Do you want to clear the textboxes?", QtGui.QMessageBox.Yes | 
            QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        
        if self.messagebox == QtGui.QMessageBox.Yes:
            self.btn3Clicked()
            self.statuslabel.setText(' - Registration completed, all textboxes cleared.')
        else:
            pass
        
        

        
    def btn3Clicked(self):
        self.le_StartTime.clear()   
        self.le_Duration.clear()
        self.le_Name.clear()
        self.le_Company.clear()
        self.le_Phone.clear()
        self.le_Products.clear()
        self.le_EquipmentNo.clear()
        self.te_Problem.clear()
        self.te_Solution.clear()
        self.cb_Solved.setCheckState(0)
        self.cb_FollowUp.setCheckState(0)
        self.cb_Forwarded.setCheckState(0)
        self.statuslabel.setText(' -All textboxes cleared.')

    def get_data(self,model):
        model.setStringList(["VolkerFitzPatrick", "Morgan Sindell", "BamNuttall", "Geometris"])

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = TogglCopy()
    sys.exit(app.exec_())


#===============================================================================
# if __name__ == '__main__':
#     main()
#===============================================================================

def btn4Clicked():
    con = lite.connect('C:\\python\\database\\test_X.db')
    cur = con.cursor()    
    cur.execute("SELECT * FROM Support")
    rows = cur.fetchall()
    s = '<html><body> <table width = "80%" align="center" border="1px" bordercolor="WhiteSmoke" ><tr><td colspan="5" bgcolor="LightGrey"><b>Overview</b></td>'
    
    
    for item in rows:
       
        if checkPhoneNumber('07590328769', item[6]):
            s += '<tr><td bgcolor="WhiteSmoke" colspan="5">Date: '+item[1][:16]+'</td></tr><tr>'
            s += '<td>'+item[4]+'</td><td>'+item[6]+'</td><td>'+item[7]+'</td><td>'+item[9]+'</td><td>'+item[10]+'</td>'
            s += '</tr><tr><td colspan="5"><hr></td></tr>'
            
    s += '</table></body></html>'
    con.close()
    htmlfile = open('C:\\python\\database\\output.html','w',encoding='utf-8')
    htmlfile.write(s)

def checkPhoneNumber(newn, oldn):
    no1 = str(newn)
    no2 = str(oldn)
    if len(no2) < 2:
        return False
    
    elif no1 == no2:
        return True
    elif no1 in no2:
        return True
    elif checkDigitForDigit(no1,no2):
        return True
    return False

def checkDigitForDigit(no1,no2):
    pattern = ''
    
    if len(no1) == len(no2):
        for i,s in enumerate(no2):
            if s == '.':
                pattern += s
            else:
                pattern += no1[i]
    
              
    
        if re.match(pattern,no2) != None:

            return True
    
    return False
    
            

print(btn4Clicked())