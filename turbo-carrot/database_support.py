# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 16:42:41 2015

@author: gjermund.vingerhagen
"""

import sys
from PyQt4 import QtGui,QtCore
import sqlite3 as lite
from datetime import datetime as dt
import re
import webbrowser
import win32com.client as win32


class TogglCopy(QtGui.QWidget):
    
    def __init__(self):
        super(TogglCopy, self).__init__()
        self.initUI()
        self.btn1Flag = False
        self.startTime = 0
    
    def initUI(self):
        self._DATABASENAME = 'N:\\Gjermund\\database_support\\support.db'
        self._USER = "GV"
        self.CLEAR_OK = True
        self.REGISTER_OK = False
        
        #Fonts
        font = QtGui.QFont( "Consolas", 12)
        # Labels
        self.label1 = QtGui.QLabel('Register case',self)
        self.label1.setFont(font)
        self.label1.move(15,10)
        
        self.label_user = QtGui.QLabel('User: {}'.format(self._USER),self)
        self.label_user.move(290,10)
        
        self.label2 = QtGui.QLabel('Start Time',self)
        self.label2.move(130,50)

        self.label3 = QtGui.QLabel('Duration',self)
        self.label3.move(230,50)
        
        self.label4 = QtGui.QLabel('Name',self)
        self.label4.move(30,100)

        self.label5 = QtGui.QLabel('Company',self)
        self.label5.move(130,100)
            
        self.label6 = QtGui.QLabel('Phone (0)   ',self)
        self.label6.move(230,100)
        
        self.label7 = QtGui.QLabel('Products',self)
        self.label7.move(30,155)

        self.label8 = QtGui.QLabel('Ser/Equip no.',self)
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
        self.btn1 = QtGui.QPushButton('', self)
        self.btn1.move(30, 35)
        self.btn1.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\clock1.png'))
        self.btn1.setIconSize(QtCore.QSize(90,55))
        self.btn1.setStyleSheet("border: 0")  
        self.btn1.clicked.connect(self.btn1Clicked)
        
        self.btn2 = QtGui.QPushButton('', self)
        self.btn2.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\register.png'))
        self.btn2.setIconSize(QtCore.QSize(30,30))
        self.btn2.setStyleSheet("border: 0")  
        self.btn2.move(30, 430)
        self.btn2.clicked.connect(self.btn2Clicked)
        
        self.btn3 = QtGui.QPushButton('', self)
        self.btn3.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\clear.png'))
        self.btn3.setIconSize(QtCore.QSize(30,30))
        self.btn3.setStyleSheet("border: 0")  
        self.btn3.move(70, 430)
        self.btn3.clicked.connect(self.btn3Clicked)
        
        self.btn4 = QtGui.QPushButton('?', self)
        self.btn4.setFixedWidth(20)
        self.btn4.move(310, 113)
        self.btn4.clicked.connect(self.btn4Clicked)
        
        self.btn_email = QtGui.QPushButton('',self)
        self.btn_email.move(255, 430)
        self.btn_email.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\email.png'))
        self.btn_email.setIconSize(QtCore.QSize(30,30))
        self.btn_email.setStyleSheet("border: 0")
        self.btn_email.clicked.connect(self.create_email)

        self.btn_task = QtGui.QPushButton('',self)
        self.btn_task.move(290, 430)
        self.btn_task.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\task.png'))
        self.btn_task.setIconSize(QtCore.QSize(30,30))
        self.btn_task.setStyleSheet("border: 0")
        self.btn_task.clicked.connect(self.create_task)

        
        #Status bar
        self.statuslabel = QtGui.QLabel('- Ready (Database = {})'.format(self._DATABASENAME),self)
        self.statuslabel.move(2,470)
        self.statuslabel.setFixedSize(341,22)
        self.statuslabel.setStyleSheet("background-color:#F8F8F8; color:grey;border: 1px solid grey; border-radius: 3px; padding: 0 4px;")
        
        #Main window
        self.setGeometry(500,300,345,493)
        self.setWindowTitle('Time registration')
        self.show()
        
        # Create an auto fill for companies
        self.completer = QtGui.QCompleter()
        self.le_Company.setCompleter(self.completer)
        model = QtGui.QStringListModel()
        self.completer.setModel(model)
        self.get_company_names(model)
        
        # Create an auto fill for phonenumbers
        self.completer2 = QtGui.QCompleter()
        self.le_Phone.setCompleter(self.completer2)
        model2 = QtGui.QStringListModel()
        self.completer2.setModel(model2)
        self.get_all_phonenumbers(model2)
        
        QtCore.QObject.connect(self.le_Phone, QtCore.SIGNAL('editingFinished()'), self.getPhoneNumberLength)
        
    def getPhoneNumberLength(self):
        l = len(self.le_Phone.text())
        st = 'Phone ({})'.format(l)
        self.label6.setText(st)
        
    def btn1Clicked(self):
        if not self.btn1Flag:
            self.startTime = dt.now()
            self.le_StartTime.setText(str(self.startTime.strftime("%H:%M")))
            self.btn1.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\clock2.png'))
        
            self.btn1.setStyleSheet("border: 0")  
            self.btn1Flag = True
            self.REGISTER_OK = False
            self.CLEAR_OK = False
            
        else:
            durtime = str((dt.now()-self.startTime)).split('.')[0]
            self.le_Duration.setText(durtime)
            self.btn1.setIcon(QtGui.QIcon('N:\Gjermund\database_support\\img\\clock1.png'))
            self.btn1Flag = False
            self.REGISTER_OK = True
                
    def btn2Clicked(self):
        if self.REGISTER_OK:
            try:
                self.con = lite.connect(self._DATABASENAME)
                
                startTime = self.startTime       
                duration = self.le_Duration.text().replace("'","")
                name = self.le_Name.text().replace("'","")
                company = self.le_Company.text().replace("'","")
                phone = self.le_Phone.text()
                products = self.le_Products.text().replace("'","")
                equipmentno = self.le_EquipmentNo.text().replace("'","")
                problem = self.te_Problem.toPlainText().replace("'","")
                solution = self.te_Solution.toPlainText().replace("'","")
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
                exestring += "'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')".format(startTime, duration, regtime, name, company, phone, products, equipmentno, problem, solution,solved,followup,forwarded,self._USER)
                print(exestring)
                with self.con:
                    cur = self.con.cursor()        
                    cur.execute(exestring)
                        
                self.con.commit()
                self.con.close()
                
                self.statuslabel.setText('- Successfully stored in database.')
                self.CLEAR_OK = True
                # Open a messagebox to potentially delete all the textboxes
                self.messagebox =  QtGui.QMessageBox.question(self, 'Message',
                    "Do you want to clear the textboxes?", QtGui.QMessageBox.Yes | 
                    QtGui.QMessageBox.No, QtGui.QMessageBox.No)
                
                if self.messagebox == QtGui.QMessageBox.Yes:
                    self.btn3Clicked()
                    self.statuslabel.setText('- Registration completed, all textboxes cleared.')
                else:
                    pass
                
            except:
                QtGui.QMessageBox.warning(self, 'Warning', 'Error code 002:\nThe registration failed. Please check your data and try again', buttons=QtGui.QMessageBox.Ok)
        
        else:
            reply = QtGui.QMessageBox.question(self, 'Continue', 'Do you want to continue without stopping the time?', buttons=QtGui.QMessageBox.Yes, defaultButton=QtGui.QMessageBox.No)    
            if reply == QtGui.QMessageBox.Yes:
                self.REGISTER_OK = True
                self.btn2Clicked()
            else:
                pass
                
    def btn3Clicked(self):
        if self.CLEAR_OK:
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
            self.statuslabel.setText('- All textboxes cleared.')
        else:
            reply = QtGui.QMessageBox.question(self, 'Continue', 'Are you sure you want to clear all field \nbefore you have registered the data?', buttons=QtGui.QMessageBox.Yes, defaultButton=QtGui.QMessageBox.No)    
            if reply == QtGui.QMessageBox.Yes:
                self.CLEAR_OK = True
                self.btn3Clicked()
            else:
                pass
            
    def get_company_names(self,model):
        model.setStringList(["Amey","VolkerFitzPatrick", "Morgan Sindell", "BamNuttall", "Geometris"])
    
    def get_all_phonenumbers(self,model2):
        try:
            con = lite.connect(self._DATABASENAME)
            cur = con.cursor()    
            cur.execute("SELECT Phone FROM Support")
            rows = cur.fetchall()        
            con.close()
            phoneset = set()
            
            for item in rows:
                if len(item[0]) == 11 and '.' not in item[0]:
                    phoneset.add(item[0])
                    
            print('Phonelist',len(phoneset))
            model2.setStringList(list(phoneset))
        except:
            QtGui.QMessageBox.warning(self, 'Warning', 'Error code 004:\nSomething went terribly wrong.', buttons=QtGui.QMessageBox.Ok)
        
    def btn4Clicked(self):
        try:
            con = lite.connect(self._DATABASENAME)
            cur = con.cursor()    
            cur.execute("SELECT * FROM Support")
            rows = cur.fetchall()
            s = '<html><body> <table width = "80%" align="center" border="1px" bordercolor="WhiteSmoke" ><tr><td colspan="5" bgcolor="LightGrey"><b>Overview</b></td>'
            
            for item in rows:
               
                if self.checkPhoneNumber(self.le_Phone.text(), item[6]):
                    s += '<tr><td bgcolor="WhiteSmoke" colspan="5">Date: '+item[1][:16]+'</td></tr><tr>'
                    s += '<td>'+item[4]+'</td><td>'+item[6]+'</td><td>'+item[7]+'</td><td>'+item[9]+'</td><td>'+item[10]+'</td>'
                    s += '</tr><tr><td colspan="5"><hr></td></tr>'
                    
            s += '</table></body></html>'
            con.close()
            htmlfile = open('N:\Gjermund\database_support\\output.html','w',encoding='utf-8')
            htmlfile.write(s)
            htmlfile.close()
            self.statuslabel.setText('- Successfully created a webpage.')
            webbrowser.open('N:\Gjermund\database_support\\output.html')
        except:
            QtGui.QMessageBox.warning(self, 'Warning', 'Error code 001\nSomething went seriously wrong', buttons=QtGui.QMessageBox.Ok)
            
    def checkPhoneNumber(self,newn, oldn):
        no1 = str(newn)
        no2 = str(oldn)
        if len(no2) < 2:
            return False
        
        elif no1 == no2:
            return True
        elif no1 in no2:
            return True
        elif self.checkDigitForDigit(no1,no2):
            return True
        return False

    def checkDigitForDigit(self,no1,no2):
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
    
    def create_email(self):
        self.cb_Forwarded.setCheckState(2)
        try:
            startTime = self.startTime       
            duration = self.le_Duration.text().replace("'","")
            name = self.le_Name.text().replace("'","")
            company = self.le_Company.text().replace("'","")
            phone = self.le_Phone.text()
            products = self.le_Products.text().replace("'","")
            equipmentno = self.le_EquipmentNo.text().replace("'","")
            problem = self.te_Problem.toPlainText().replace("'","")
            solution = self.te_Solution.toPlainText().replace("'","")
            
            
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.Subject = '{} ({}) - Phone: {}'.format(name, company,phone)
            mail.Body = """Instrument: {}
Serial/Equipmentnr: {}
       
{}

{}

Kind regards,
{}   
""".format(products,equipmentno,problem,solution,self._USER)
            mail.Display(True)
            self.cb_Forwarded.checkState(1)
        except:
            QtGui.QMessageBox.warning(self, 'Warning', 'Email could not be sent', buttons=QtGui.QMessageBox.Ok)

    def create_task(self):
        self.cb_FollowUp.setCheckState(2)
        try:
            startTime = self.startTime       
            duration = self.le_Duration.text().replace("'","")
            name = self.le_Name.text().replace("'","")
            company = self.le_Company.text().replace("'","")
            phone = self.le_Phone.text()
            products = self.le_Products.text().replace("'","")
            equipmentno = self.le_EquipmentNo.text().replace("'","")
            problem = self.te_Problem.toPlainText().replace("'","")
            solution = self.te_Solution.toPlainText().replace("'","")
            
            
            outlook = win32.Dispatch('outlook.application')
            task = outlook.CreateItem(3)
            task.Subject = '{} ({}) - Phone: {}'.format(name, company,phone)
            task.Body = """Starttime: {}
            
Instrument: {}
Serial/Equipmentnr: {}
       
{}

{}""".format(startTime,products,equipmentno,problem,solution)
            task.Display(True)
            
        except:
            QtGui.QMessageBox.warning(self, 'Warning', 'Task could not be created', buttons=QtGui.QMessageBox.Ok)

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = TogglCopy()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()