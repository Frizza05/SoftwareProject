# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'LittleSisterGUI.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui
import qdarkstyle
import sys
import csv
import os
import pandas as pd
from datetime import datetime
import shutil
from qrtools import QR
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

# Starting window for program
class Ui_MainWindow(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self)
    def setupUi(self, MainWindow):
#defining the layout of the window
        MainWindow.setObjectName(_fromUtf8("Shift Manager"))
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.gridLayout = QtGui.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))

#Label - Program Title (Logo)
        self.labelTitle = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(24)
        self.labelTitle.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.labelTitle.setAlignment(QtCore.Qt.AlignCenter)
        self.labelTitle.setObjectName(_fromUtf8("labelTitle"))
        self.labelTitle.setMinimumSize(QtCore.QSize(700, 0))
        self.gridLayout.addWidget(self.labelTitle, 0, 0, 1, 1)
        self.horizontalLayout_3 = QtGui.QHBoxLayout()
        self.horizontalLayout_3.setObjectName(_fromUtf8("horizontalLayout_3"))

#LCD Timer - Displays Exact Time
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.Time)
        self.timer.start(1000)
        self.lcd = QtGui.QLCDNumber(self)
#Initially displays current time in Hour:Minute format
        self.lcd.display(datetime.now().strftime("%H"+":"+"%M"))
        self.lcd.setMinimumSize(QtCore.QSize(0,100))
        self.horizontalLayout_3.addWidget(self.lcd)
        
#Label - Welcome ID and Program Messaging
        self.gridLayout.addLayout(self.horizontalLayout_3, 1, 0, 1, 1)
        self.label = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(24)
        self.label.setFont(font)
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName(_fromUtf8("label"))
        self.gridLayout.addWidget(self.label, 2, 0, 1, 1)
        spacerItem = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Minimum)

        self.gridLayout.addItem(spacerItem, 3, 0, 1, 1)
        self.horizontalLayout = QtGui.QHBoxLayout()
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.horizontalLayout_5 = QtGui.QHBoxLayout()
        self.horizontalLayout_5.setObjectName(_fromUtf8("horizontalLayout_5"))

#Push Button - Start the QR code scanning
        self.buttonStartScan = QtGui.QPushButton(self.centralwidget)
        self.buttonStartScan.setMinimumSize(QtCore.QSize(0, 200))
        self.buttonStartScan.setObjectName(_fromUtf8("buttonStartScan"))
        self.horizontalLayout_5.addWidget(self.buttonStartScan)
        self.horizontalLayout.addLayout(self.horizontalLayout_5)
        self.buttonStartScan.setStyleSheet("background-color:rgb(49,54,59)")
        self.verticalLayout = QtGui.QVBoxLayout()
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))

#Push Button - Start Employee's shift
        self.buttonStartShift = QtGui.QPushButton(self.centralwidget)
        self.buttonStartShift.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonStartShift.setObjectName(_fromUtf8("buttonStartShift"))
        self.verticalLayout.addWidget(self.buttonStartShift)
        self.buttonStartShift.setStyleSheet("background-color:rgb(49,54,59)")

#Push Button - Finish Employee's Shift
        self.buttonFinishShift = QtGui.QPushButton(self.centralwidget)
        self.buttonFinishShift.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonFinishShift.setObjectName(_fromUtf8("buttonFinishShift"))
        self.verticalLayout.addWidget(self.buttonFinishShift)
        self.buttonFinishShift.setStyleSheet("background-color:rgb(49,54,59)")
        
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.gridLayout.addLayout(self.horizontalLayout, 4, 0, 1, 1)
        self.horizontalLayout_2 = QtGui.QHBoxLayout()
        self.horizontalLayout_2.setObjectName(_fromUtf8("horizontalLayout_2"))
        self.verticalLayout_3 = QtGui.QVBoxLayout()
        self.verticalLayout_3.setObjectName(_fromUtf8("verticalLayout_3"))
        spacerItem1 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Minimum)
        self.verticalLayout_3.addItem(spacerItem1)
        self.horizontalLayout_4 = QtGui.QHBoxLayout()
        self.horizontalLayout_4.setObjectName(_fromUtf8("horizontalLayout_4"))

#Push Button - Access Manager Options  
        self.buttonManagerOptions = QtGui.QPushButton(self.centralwidget)
        self.buttonManagerOptions.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonManagerOptions.setObjectName(_fromUtf8("pushButton"))
        self.horizontalLayout_4.addWidget(self.buttonManagerOptions)
        self.buttonManagerOptions.setStyleSheet("background-color:rgb(49,54,59)")

#Push Button - Employee logout   
        self.buttonLogOut = QtGui.QPushButton(self.centralwidget)
        self.buttonLogOut.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonLogOut.setObjectName(_fromUtf8("buttonLogOut"))
        self.horizontalLayout_4.addWidget(self.buttonLogOut)
        self.buttonLogOut.setStyleSheet("background-color:rgb(49,54,59)")

#Push Button - Exit the whole program  
        self.buttonExit = QtGui.QPushButton(self.centralwidget)
        self.buttonExit.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonExit.setObjectName(_fromUtf8("buttonExit"))
        self.horizontalLayout_4.addWidget(self.buttonExit)
        self.buttonExit.setStyleSheet("background-color:rgb(189,0,0)")
        
        self.verticalLayout_3.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_2.addLayout(self.verticalLayout_3)
        self.gridLayout.addLayout(self.horizontalLayout_2, 5, 0, 1, 1)
        self.setLayout(self.gridLayout)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

#Global variable for ID initially setting it to 0
        global ID
        ID = 0

    def retranslateUi(self, MainWindow):

        MainWindow.setWindowTitle(_translate("MainWindow", "Shift Manager", None))
        #Locating Little Sister Logo and displaying it as the label - Title
        pixmap = QtGui.QPixmap('littlesisterlogo.png') 
        self.labelTitle.setPixmap(pixmap)
    
        self.label.setText(_translate("MainWindow", "Please Scan QR Code", None))
        self.buttonStartScan.setText(_translate("MainWindow", "Scan QR Code", None))
        self.buttonStartShift.setText(_translate("MainWindow", "Start Shift", None))
        self.buttonFinishShift.setText(_translate("MainWindow", "Finish Shift", None))
        self.buttonManagerOptions.setText(_translate("MainWindow", "Manager Options", None))
        self.buttonLogOut.setText(_translate("MainWindow", "Log Out", None))
        self.buttonExit.setText(_translate("MainWindow", "Exit", None))

        #Connecting all the buttons to function when clicked
        self.buttonStartShift.clicked.connect(self.StartShift)
        self.buttonStartScan.clicked.connect(self.StartScan)
        self.buttonFinishShift.clicked.connect(self.FinishShift)
        self.buttonExit.clicked.connect(self.Exit)
        self.buttonLogOut.clicked.connect(self.LogOut)
        self.buttonManagerOptions.clicked.connect(self.ManagerOptions)

    #Whenever the LCD calls for the time
    def Time(self):
        self.lcd.display(datetime.now().strftime("%H"+":"+"%M"))

    #When the button "Start Scan" is pressed
    def StartScan(self):
        #Reading the file "EmployeeList.csv" as a dataframe
        data = pd.read_csv("EmployeeList.csv")
        #Defining global variables so other functions can access them
        global name
        option = 0
        name = 0
        ID_found = False
        count = 0
        #Setting the variable "myCode" to access the function QR and Decode webcam images
        myCode = QR()
        myCode.decode_webcam()
        global ID
        global manager
        #Setting variable "ID" to equal the scanned code
        ID = myCode.data
        #Searching the file "EmployeeList.csv" to see if the ID exists, then grabbing the fullname and role
        #Also changing the label name say "Successfully Scanned" for 3 seconds when the ID is found
        for i in range((len(data["ID"]))):
            if data["ID"][i] == ID:
                name = (data["Name"][i])
                ID_found = True
                manager = (data["Role"][i])
                self.label.setStyleSheet('color: rgb(0, 183, 0)')
                self.label.setText("Successfully Scanned")
                QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Welcome '+name))
                QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
                break
            else:
                count = count + 1
        #If the ID cannot be found within the file the label changes
        if ID_found == False:
            self.label.setText("ID Not Found, Please Try Again")
    #When the button "Start Shift" is pressed       
    def StartShift(self):
        #If the ID = 0 (ID is not scanned) a message box will appear
        if ID == 0:
            QtGui.QMessageBox.warning(self, "ERROR","Please Scan QR Code")
        else:
            #See if file is there and if not, create it
            try:
                with open('EmployeeIN.csv') as file:
                   reader = csv.reader(file)
                   file.close()
            except IOError:
                #Create the file and write the correct headers
                with open('EmployeeIN.csv', 'ab') as csvf:
                    writer = csv.writer(csvf)
                    writer.writerow(["ID", "Date", "Day", "In", "Out", "Total"])
                    csvf.close()
            try:
                #Open File and see if current login is unique (havent logged in currently)
                with open('EmployeeIN.csv', 'rb') as f:
                    reader = csv.reader(f)
                    unique = True
                    #defining current date
                    date = datetime.now().strftime("%d-%m-%y")

                    #Searching each row to see if unique
                    for row in reader:
                        if row[0] == ID and row[1] == date:
                            unique = False
                    f.close()

                #Open "EmployeeIN.csv" file to append row with starting shift details
                #If it is a unique entry then append it to the bottom of the .csv
                with open('EmployeeIN.csv', 'ab') as f:
                    writer = csv.writer(f)
                    
                    if unique == True:
                        writer.writerow([ID,datetime.now().strftime("%d-%m-%y"),datetime.strptime(date,'%d-%m-%y').strftime('%A'),datetime.now().strftime("%H:%M")])
                        self.label.setStyleSheet('color: rgb(0, 183, 0)')
                        self.label.setText("Shift Started")
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Welcome '+name))
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
                        
                    if unique == False:
                        self.label.setStyleSheet('color: rgb(189, 0, 0)')
                        self.label.setText("Shift Already Started")
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Welcome '+name))
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
                    f.close()
                    
            #If the ID cannot be located label is changed to print "ID Not Found"
            except NameError:
                self.label.setStyleSheet('color: rgb(189,0,0)')
                self.label.setText("ID Not Found")
                QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Welcome '+name))
                QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
    #When the button "Finish Shift" is pressed
    def FinishShift(self):
        #If the employee hasnt scanned in message box is shown, else continue on
        if ID == 0:
            QtGui.QMessageBox.warning(self, "ERROR","Please Scan QR Code")
        else:
            date = datetime.now().strftime("%d-%m-%y")
            time = datetime.now().strftime("%H:%M")
            t = "%H:%M"
            counter = 0
            found = False
            
            #Open "EmployeeIN.csv" file for read to see if employee has correctly started their shift
            with open('EmployeeIN.csv', 'rb') as csvfile:
                reader = csv.reader(csvfile)
                employdata = []
                while found == False:
                    for row in reader:
                        #If the first item of the row = the correct ID and the second = current date
                        if row[0] == ID and row[1] == date:
                            timein = row[3]
                            x = 0
                            found = True
                            #Appending each item in the row to a list "employdata"
                            for x in range(len(row)):
                               employdata.append(row[x])
                    
                    #If the shift cannot be found then print "Shift Not Started"
                    if found == False:
                        self.label.setStyleSheet('color: rgb(189, 0, 0)')
                        self.label.setText("Shift Not Started")
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Welcome '+name))
                        QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
                    csvfile.close()
            
                #Add shift finishing time to list
                employdata.append(time)
                #Calculate the total time (Timefinal-Timeinitial) and convert into string
                timetotal = str(datetime.strptime(time,t)-datetime.strptime(timein,t))
                #Strip last 3 characters so it is formated 00:00
                timetotal = timetotal[:-3]
                #Add total time to list
                employdata.append(timetotal)
                    
            #Try open "EmployeeOUT.csv" for read to see if it exists
            try:
                with open('EmployeeOUT.csv','rb') as file:
                    reader = csv.reader(file)
                    file.close()
            #If the file doesnt exist, then create it with the correct headers
            except IOError:
                with open('EmployeeOUT.csv','ab') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["ID","Date","Day","In","Out","Total"])
                    csvfile.close()
            #Append the list "employlist" to the "EmployeeOUT.csv" file which is the employees data               
            with open('EmployeeOUT.csv', 'ab') as f:
                writer = csv.writer(f)
                writer.writerow(employdata)
                self.label.setText("Shift Ended")
                f.close()

            #To remove the row from the "EmployeeIN.csv" a new file as to be created
            #writing all the lines except the one we dont need essentially deleting the row
            with open('EmployeeIN.csv', 'rb') as inp, open('EmployeeIN2.csv', 'ab') as out:
                writer = csv.writer(out)
                for row in csv.reader(inp):
                    if row[0] != ID:
                        writer.writerow(row)
                inp.close()
                out.close()
            #Removing the old file and replacing it with the new one
            os.remove('EmployeeIN.csv')
            os.rename('EmployeeIN2.csv','EmployeeIN.csv')
    #When the button "Exit" is pressed close the program
    def Exit(self):
        sys.exit()
    #When the button "Log Out" is pressed
    def LogOut(self):
        #Changing all employee variables to 0 essentially logging them out
        global ID
        ID = 0
        global manager
        manager = 0
        #Displaying a successful message as the label when the program has logged them out
        self.label.setStyleSheet('color: rgb(0, 183, 0)')
        self.label.setText("Successfully Logged Out")
        QtCore.QTimer.singleShot(3000, lambda: self.label.setText('Please Scan QR Code'))
        QtCore.QTimer.singleShot(3000, lambda: self.label.setStyleSheet('color: white'))
    #When the button "Manager Options" is pressed       
    def ManagerOptions(self):
        #Checking to see if they have logged in and also have the right permissions
        if ID == 0:
            QtGui.QMessageBox.warning(self, "ERROR","Please Scan QR Code")
        elif manager == 0:
            QtGui.QMessageBox.critical(self,"Permission Denied", "Employee Does Not Have Manager Privileges")
        elif manager == 1:
            #closing the current windows and showing the manager options
            ex2.show()
            self.close()
            
#Manager Options Window   
class Ui_ManagerOptionsWindow(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self) 
    def setupUi(self, ManagerOptionsWindow):
        ManagerOptionsWindow.setObjectName(_fromUtf8("ManagerOptionsWindow"))
        self.centralwidget = QtGui.QWidget(ManagerOptionsWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.gridLayout = QtGui.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        
        #Push Button - Open new window to print files data
        self.buttonPrintTimes = QtGui.QPushButton(self.centralwidget)
        self.buttonPrintTimes.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonPrintTimes.setObjectName(_fromUtf8("buttonPrintTimes"))
        self.gridLayout.addWidget(self.buttonPrintTimes, 4, 0, 1, 1)
        self.buttonPrintTimes.setStyleSheet("background-color:rgb(49,54,59)")

        #Push Button - Open new window to create new employee
        self.buttonAddEmployee = QtGui.QPushButton(self.centralwidget)
        self.buttonAddEmployee.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonAddEmployee.setObjectName(_fromUtf8("buttonAddEmployee"))
        self.gridLayout.addWidget(self.buttonAddEmployee, 3, 0, 1, 1)
        self.buttonAddEmployee.setStyleSheet("background-color:rgb(49,54,59)")

        #Push Button - Finish all current open shifts
        self.buttonFinishShifts = QtGui.QPushButton(self.centralwidget)
        self.buttonFinishShifts.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonFinishShifts.setObjectName(_fromUtf8("buttonFinishShifts"))
        self.gridLayout.addWidget(self.buttonFinishShifts, 5, 0, 1, 1)
        self.buttonFinishShifts.setStyleSheet("background-color:rgb(49,54,59)")

        #Push Button - Go back to previous window
        self.buttonGoBack = QtGui.QPushButton(self.centralwidget)
        self.buttonGoBack.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonGoBack.setObjectName(_fromUtf8("buttonGoBack"))
        self.gridLayout.addWidget(self.buttonGoBack, 9, 0, 1, 1)
        self.buttonGoBack.setStyleSheet("background-color:rgb(189,0,0)")

        #Push Button - Export the "EmployeeOUT.csv" file via email
        self.buttonExportTimes = QtGui.QPushButton(self.centralwidget)
        self.buttonExportTimes.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonExportTimes.setObjectName(_fromUtf8("buttonExportTimes"))
        self.gridLayout.addWidget(self.buttonExportTimes, 6, 0, 1, 1)
        self.buttonExportTimes.setStyleSheet("background-color:rgb(49,54,59)")

        #Spacer
        spacerItem = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem, 2, 0, 1, 1)
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setMinimumSize(QtCore.QSize(0, 60))
        self.label.setMaximumSize(QtCore.QSize(16777215, 60))

        #Label - Title
        font = QtGui.QFont()
        font.setPointSize(30)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName(_fromUtf8("label"))
        self.label.setMinimumSize(QtCore.QSize(650, 0))
        
        self.gridLayout.addWidget(self.label, 1, 0, 1, 1)
        spacerItem1 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem1, 7, 0, 1, 1)
        self.retranslateUi(ManagerOptionsWindow)
        QtCore.QMetaObject.connectSlotsByName(ManagerOptionsWindow)
        self.setLayout(self.gridLayout)

        #Progress Bar - Shows progress and initially set to a value of 0
        self.progressBar = QtGui.QProgressBar(self.centralwidget)
        self.gridLayout.addWidget(self.progressBar, 8,0,1,1)
        self.progressBar.setValue(0)
                
    def retranslateUi(self, ManagerOptionsWindow):
        ManagerOptionsWindow.setWindowTitle(_translate("ManagerOptionsWindow", "Manager Options", None))
        self.buttonFinishShifts.setText(_translate("ManagerOptionsWindow", "Finish All Employee Shifts", None))
        self.buttonPrintTimes.setText(_translate("ManagerOptionsWindow", "Print Employee Times", None))
        self.buttonAddEmployee.setText(_translate("ManagerOptionsWindow", "Add Employee", None))
        self.buttonGoBack.setText(_translate("ManagerOptionsWindow", "Go Back", None))
        self.buttonExportTimes.setText(_translate("ManagerOptionsWindow", "Export Employee Times (Email)", None))
        self.label.setText(_translate("ManagerOptionsWindow", "Manager Options", None))       

        #Connecting all the buttons to function when clicked
        self.buttonGoBack.clicked.connect(self.GoBack)
        self.buttonAddEmployee.clicked.connect(self.AddEmployee)
        self.buttonExportTimes.clicked.connect(self.ExportTimes)
        self.buttonPrintTimes.clicked.connect(self.PrintTimes)
        self.buttonFinishShifts.clicked.connect(self.FinishShifts)

    #When the button "Add Employee" is pressed   
    def AddEmployee(self):
        #Show the "Add Employee" window and close the current window
        addEmployeeWindow.show()
        self.close()
        
    #When the button "Go Back" is pressed  
    def GoBack(self):
        #Show the previous homepage window and close the current window
        ex.show()
        self.close()

    #When the button "Export Employee Times" is pressed
    def ExportTimes(self):
        #Set progress bar to 0 and start email process
        self.progressBar.setValue(0)

        #To and from email addresses (Constant)
        fromaddr = "frisbe05@live.com"
        toaddr = "friswk13@sbc.vic.edu.au"
        #Increase in progress bar
        self.progressBar.setValue(12.5)
        
        msg = MIMEMultipart() 
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = "Little Sister Employee Times"
        #Increase in progress bar
        self.progressBar.setValue(25)

        body = "Attached is this weeks employee's times."
        msg.attach(MIMEText(body, 'plain')) 
        filename = "EmployeeOUT.csv"
        attachment = open("EmployeeOUT.csv", "rb")
        #Increase in progress bar
        self.progressBar.setValue(37.5)
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        #Increase in progress bar
        self.progressBar.setValue(50)

        msg.attach(part)
        server = smtplib.SMTP('smtp.live.com', 587)
        #Increase in progress bar
        self.progressBar.setValue(62.5)
        
        server.starttls()
        server.login(fromaddr, "lala05")
        #Increase in progress bar
        self.progressBar.setValue(75)

        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        #Increase in progress bar
        self.progressBar.setValue(87.5)

        #Move the exported file to another folder for future storage
        folder = os.getcwd()+"/OldExports/"
        shutil.move(os.getcwd()+"/EmployeeOUT.csv",folder+"Exported_"+datetime.now().strftime("%d-%m-%y")+".csv")

        #Finish Progress Bar
        self.progressBar.setValue(100)

    #When the button "Finish All Shifts" is pressed
    def FinishShifts(self):
        #Set progress bar to 0
        self.progressBar.setValue(0)
        #Open "EmployeeIN.csv" as dataframe
        data = pd.read_csv("EmployeeIN.csv")

        #Check to see if file has no current employees with started shifts
        if len(data["ID"]) == 0:
            self.label.setStyleSheet('color: rgb(189, 0, 0)')
            self.label.setText("No Shifts Started")
            QtCore.QTimer.singleShot(4000, lambda: self.label.setText('Manager Options'))
            QtCore.QTimer.singleShot(4000, lambda: self.label.setStyleSheet('color: white'))

        #If there is at least 1 employee with a started shift
        elif len(data["ID"]) != 0:
            #Set progress bar to 0
            self.progressBar.setValue(50) 
            counter = 0
            #Open "EmployeeIN.csv" for reading
            with open('EmployeeIN.csv','rb') as f:
                reader = csv.reader(f)
                #for each row in the reader that isnt the header copy the data and append it to "EmployeeOUT.csv"
                for row in reader:
                    employeedata = []
                    if row[0] != "ID":
                        x = 0
                        for x in range(len(row)):
                            employeedata.append(row[x])
                        #Write last to columns in employees shift as "N/A"
                        for i in range(2):
                            employeedata.append("N/A")
                        with open('EmployeeOUT.csv','ab') as csvfile:
                            writer = csv.writer(csvfile)
                            writer.writerow(employeedata)
            f.close()
            #Update progress bar
            self.progressBar.setValue(50)

            #Open "EmployeeIN.csv", overwriting current data (essentially deleting all data)
            with open('EmployeeIN.csv','w+') as file:
                writer = csv.writer(file)
                writer.writerow(["ID","Date","Day","In","Out","Total"])
                f.close()
            #Confirmation to the user that it was successful, changing the label color and text
            self.label.setStyleSheet('color: rgb(0, 189, 0)')
            self.label.setText("Successfully Finished Shifts")
            QtCore.QTimer.singleShot(4000, lambda: self.label.setText('Manager Options'))
            QtCore.QTimer.singleShot(4000, lambda: self.label.setStyleSheet('color: white'))
            #Finishing progress bar
            self.progressBar.setValue(100)
    #When the button "Print Times" is pressed    
    def PrintTimes(self):
        #Open the new windows and close the current window
        self.close()
        printCSV.show()
#New Employee Window
class Ui_AddEmployee(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self)
    def setupUi(self, AddEmployee):
        AddEmployee.setObjectName(_fromUtf8("AddEmployee"))
        AddEmployee.resize(410, 480)
        self.centralwidget = QtGui.QWidget(AddEmployee)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.verticalLayout = QtGui.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))

        #Label - Title
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setMinimumSize(QtCore.QSize(0, 60))
        self.label.setMaximumSize(QtCore.QSize(16777215, 60))
        font = QtGui.QFont()
        font.setPointSize(30)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName(_fromUtf8("label"))
        self.verticalLayout.addWidget(self.label)

        #Label - First Name input
        self.horizontalLayout = QtGui.QHBoxLayout()
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        self.labelFirstName = QtGui.QLabel(self.centralwidget)
        self.labelFirstName.setObjectName(_fromUtf8("labelFirstName"))
        self.horizontalLayout.addWidget(self.labelFirstName)

        #Text Edit - Input First Name
        self.lineEdit = QtGui.QLineEdit(self.centralwidget)
        self.lineEdit.setMinimumSize(QtCore.QSize(0, 30))
        self.lineEdit.setObjectName(_fromUtf8("lineEdit"))
        self.horizontalLayout.addWidget(self.lineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtGui.QHBoxLayout()
        self.horizontalLayout_2.setObjectName(_fromUtf8("horizontalLayout_2"))

        #Label - Last Name input
        self.labelLastName = QtGui.QLabel(self.centralwidget)
        self.labelLastName.setObjectName(_fromUtf8("labelLastName"))
        self.horizontalLayout_2.addWidget(self.labelLastName)

        #Text Edit - Input last name
        self.lineEdit_2 = QtGui.QLineEdit(self.centralwidget)
        self.lineEdit_2.setMinimumSize(QtCore.QSize(0, 30))
        self.lineEdit_2.setObjectName(_fromUtf8("lineEdit_2"))
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.verticalLayout.addLayout(self.horizontalLayout_2)

        #Spacer
        spacerItem = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout_3 = QtGui.QHBoxLayout()
        self.horizontalLayout_3.setObjectName(_fromUtf8("horizontalLayout_3"))

        #Label - Asking if new employee needs manager privileges
        self.labelManagerRole = QtGui.QLabel(self.centralwidget)
        self.labelManagerRole.setObjectName(_fromUtf8("labelManagerRole"))
        self.horizontalLayout_3.addWidget(self.labelManagerRole)
        self.verticalLayout_2 = QtGui.QVBoxLayout()
        self.verticalLayout_2.setObjectName(_fromUtf8("verticalLayout_2"))

        #Radio Button - Yes
        self.radioYes = QtGui.QRadioButton(self.centralwidget)
        self.radioYes.setObjectName(_fromUtf8("radioYes"))
        self.verticalLayout_2.addWidget(self.radioYes)

        #Radio Button - No
        self.radioNo = QtGui.QRadioButton(self.centralwidget)
        self.radioNo.setObjectName(_fromUtf8("radioNo"))
        self.verticalLayout_2.addWidget(self.radioNo)
        self.horizontalLayout_3.addLayout(self.verticalLayout_2)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_4 = QtGui.QHBoxLayout()
        self.horizontalLayout_4.setContentsMargins(-1, 0, -1, -1)
        self.horizontalLayout_4.setObjectName(_fromUtf8("horizontalLayout_4"))

        #Label - Enter Employee Email
        self.label_2 = QtGui.QLabel(self.centralwidget)
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.horizontalLayout_4.addWidget(self.label_2)

        #Text Edit - Input Employee Email
        self.lineEdit_3 = QtGui.QLineEdit(self.centralwidget)
        self.lineEdit_3.setMinimumSize(QtCore.QSize(0, 30))
        self.lineEdit_3.setObjectName(_fromUtf8("lineEdit_3"))
        self.horizontalLayout_4.addWidget(self.lineEdit_3)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        spacerItem1 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)

        #Progress bar initially set to 0
        self.progressBar = QtGui.QProgressBar(self.centralwidget)
        self.verticalLayout.addWidget(self.progressBar)
        self.progressBar.setValue(0)

        #Push Button - Submit Details
        self.buttonSubmit = QtGui.QPushButton(self.centralwidget)
        self.buttonSubmit.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonSubmit.setObjectName(_fromUtf8("buttonSubmit"))
        self.verticalLayout.addWidget(self.buttonSubmit)
        self.buttonSubmit.setStyleSheet("background-color:rgb(49,54,59)")

        #Spacer
        spacerItem2 = QtGui.QSpacerItem(20, 40, QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem2)

        #Push Button - Go Back to previous window
        self.buttonGoBack = QtGui.QPushButton(self.centralwidget)
        self.buttonGoBack.setMinimumSize(QtCore.QSize(0, 50))
        self.buttonGoBack.setObjectName(_fromUtf8("buttonGoBack"))
        self.verticalLayout.addWidget(self.buttonGoBack)
        self.buttonGoBack.setStyleSheet("background-color:rgb(189,0,0)")
        
        self.setLayout(self.verticalLayout)
        self.retranslateUi(AddEmployee)
        QtCore.QMetaObject.connectSlotsByName(AddEmployee)



    def retranslateUi(self, AddEmployee):
        AddEmployee.setWindowTitle(_translate("AddEmployee", "Add New Employee", None))
        self.label.setText(_translate("AddEmployee", "Add New Employee", None))
        self.labelFirstName.setText(_translate("AddEmployee", "First Name:", None))
        self.labelLastName.setText(_translate("AddEmployee", "Last Name:", None))
        self.labelManagerRole.setText(_translate("AddEmployee", "New Employee Require Manager Privilege?", None))
        self.radioYes.setText(_translate("AddEmployee", "Yes", None))
        self.radioNo.setText(_translate("AddEmployee", "No", None))
        self.label_2.setText(_translate("AddEmployee", "Enter Employee Email:", None))
        self.buttonSubmit.setText(_translate("AddEmployee", "Submit New Employee", None))
        self.buttonGoBack.setText(_translate("AddEmployee", "Go Back", None))

        #Connecting all the buttons to function when clicked
        self.buttonGoBack.clicked.connect(self.GoBack)
        self.buttonSubmit.clicked.connect(self.Submit)
        self.radioYes.clicked.connect(self.yesOn)
        self.radioNo.clicked.connect(self.noOn)

    def Submit(self):
        #Set Progress bar to 0
        self.progressBar.setValue(0)

        #Setting variables to the text input from the text edit
        number = 1
        first_name = str(self.lineEdit.text())
        last_name = str(self.lineEdit_2.text())
        toaddr = str(self.lineEdit_3.text())
        new_ID = str.upper("".join([last_name[:4],"00",str(number)]))
        unique = False
        #Updating progress baar
        self.progressBar.setValue(10)

        #Open "EmployeeList.csv" for reading to create a unique ID
        with open('EmployeeList.csv', 'rb') as f:
            reader = csv.reader(f)
            while unique == False:
                for row in reader:
                    #If ID is not unique change last digit by 1
                    if row[0] == new_ID:
                        number = number + 1
                        new_ID = str.upper("".join([last_name[:4],"00",str(number)]))
                unique = True

        #Update Progress bar
        self.progressBar.setValue(20)

        #Create unique QR code for employee and finding filepath for it
        myCode = QR(data=new_ID, pixel_size=10)
        myCode.encode()
        filepath = myCode.filename
        #Update progress bar
        self.progressBar.setValue(30)

        #Moving QR file from temp folder into directory folder
        QRfile = filepath.split("/")
        shutil.move(filepath,os.getcwd()+"/"+new_ID+".png")
        #Update Progress Bar
        self.progressBar.setValue(40)

        #Email QR code file to employee's email address
        fromaddr = "frisbe05@live.com"
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = 'Employee QR Code'
        body = 'Attached is your QR code'
        self.progressBar.setValue(50)
        
        msg.attach(MIMEText(body, 'plain'))
        filename = new_ID+".png"
        attachment = open(new_ID+".png", "rb")
        self.progressBar.setValue(60)
         
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        self.progressBar.setValue(70)
        
        msg.attach(part)
        server = smtplib.SMTP('smtp.live.com', 587)
        server.starttls()
        self.progressBar.setValue(80)
        
        server.login(fromaddr, "lala05")
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        self.progressBar.setValue(90)

        #Open "EmployeeList.csv" apending their information to it 
        with open('EmployeeList.csv','ab') as f:
            writer = csv.writer(f)
            writer.writerow([new_ID,(" ".join([first_name,last_name])),manager_status])
            f.close()
        #Remove newly created QR file
        os.remove(new_ID+".png")

        #Clear all text edit lines for easy input of new employee
        self.lineEdit.clear()
        self.lineEdit_2.clear()
        self.lineEdit_3.clear()
        self.radioNo.setChecked(True)
        #Finish progress bar
        self.progressBar.setValue(100)

    #When the radio toggle "Yes" is selected
    def yesOn(self):
        #Set manager status to 1 (Manager privileges)
        global manager_status
        manager_status = 1
    #When the radio toggle "No" is selected
    def noOn(self):
        #Set manager status to 0 (No manager privileges)
        global manager_status
        manager_status = 0
    #When the button "Go Back" is pressed
    def GoBack(self):
        #Close current window and open previous
        ex2.show()
        self.close()
#Displaying "EmployeeIN.csv" and "EmployeeOUT.csv" files
class Ui_PrintCSV(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self)
    def setupUi(self, PrintCSV):
        PrintCSV.setObjectName(_fromUtf8("Print Employee Times"))
        PrintCSV.setWindowTitle(_translate("PrintCSV", "Print Employee Times", None))
        PrintCSV.resize(650, 500)
        self.model = QtGui.QStandardItemModel(self)

        #Table view for showing csv files
        self.tableView = QtGui.QTableView(self)
        self.tableView.setModel(self.model)
        self.tableView.horizontalHeader().setStretchLastSection(True)

        #Push Button - Load "EmployeeIN.csv" file
        self.pushButtonLoadIN = QtGui.QPushButton(self)
        self.pushButtonLoadIN.setText("Print In CSV")
        self.pushButtonLoadIN.clicked.connect(self.loadCSVIN)
        self.pushButtonLoadIN.setMinimumSize(QtCore.QSize(0, 50))
        self.pushButtonLoadIN.setStyleSheet("background-color:rgb(49,54,59)")

        #Push Button - Go Back to previous window
        self.pushButtonExit = QtGui.QPushButton(self)
        self.pushButtonExit.setText("Go Back")
        self.pushButtonExit.clicked.connect(self.GoBack)
        self.pushButtonExit.setMinimumSize(QtCore.QSize(0, 50))
        self.pushButtonExit.setStyleSheet("background-color:rgb(189,0,0)")

        #Push Button - Load "EmployeeOUT.csv" file
        self.pushButtonLoadOUT = QtGui.QPushButton(self)
        self.pushButtonLoadOUT.setText("Print Out CSV")
        self.pushButtonLoadOUT.clicked.connect(self.loadCSVOUT)
        self.pushButtonLoadOUT.setMinimumSize(QtCore.QSize(0, 50))
        self.pushButtonLoadOUT.setStyleSheet("background-color:rgb(49,54,59)")

        #Setting layout of window as vertical 
        self.layoutVertical = QtGui.QVBoxLayout(self)
        self.layoutVertical.addWidget(self.tableView)
        self.layoutVertical.addWidget(self.pushButtonLoadIN)
        self.layoutVertical.addWidget(self.pushButtonLoadOUT)
        self.layoutVertical.addWidget(self.pushButtonExit)
        
    #When the button "Print In CSV" is pressed
    def loadCSVIN(self,filename):
        #Clear the table view
        self.model.setRowCount(0)

        #Open "EmployeeIN.csv" file for reading and copying data to list 
        with open("EmployeeIN.csv","rb") as fileinput:
            for row in csv.reader(fileinput):
                items = [
                    QtGui.QStandardItem(field)
                    for field in row
                ]
                self.model.appendRow(items)
    #When the button "Print Out CSV" is pressed
    def loadCSVOUT(self,filename):
        #Clear and data currently in the table view
        self.model.setRowCount(0)

        #Try Opening the file "EmployeeOUT.csv"
        try:
            #Open "EmployeeOUT.csv" file for reading
            with open("EmployeeOUT.csv","rb") as fileinput:
                for row in csv.reader(fileinput):
                    items = [
                        QtGui.QStandardItem(field)
                        for field in row
                    ]
                    self.model.appendRow(items)
                    
        #If the file cannot be opened (IOError) then display error message box
        except IOError:
            QtGui.QMessageBox.warning(self, "ERROR","No File Available: File Already Exported")

    #When the button "Go Back" is pressed
    def GoBack(self):
        #Close current window and open previous
        self.close()
        ex2.show()

if __name__ == "__main__":
    #Globally set variables for windows
    global app
    app = QtGui.QApplication(sys.argv)
    
    #Set Style of windows to QDarkStyle
    app.setStyleSheet(qdarkstyle.load_stylesheet(pyside=False))
    
    global ex
    ex = Ui_MainWindow()
    #Show starting window
    ex.show()
    global ex2
    ex2 = Ui_ManagerOptionsWindow()
    global addEmployeeWindow
    addEmployeeWindow = Ui_AddEmployee()
    global printCSV
    printCSV = Ui_PrintCSV()
    sys.exit(app.exec_())

