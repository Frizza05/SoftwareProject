import os
import sys
import csv
import pandas as pd
from datetime import datetime
import shutil
from qrtools import QR
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

data = pd.read_csv("EmployeeList.csv")

option = 0
name = 0
ID_found = False
manager = False

while ID_found == False:
    count = 0
    global myCode
    global scanner
    myCode = QR()
    scanner = myCode.decode_webcam()
    ID = myCode.data
    for i in range((len(data["ID"]))):
        if data["ID"][i] == ID:
            name = (data["Name"][i])
            ID_found = True
            manager = (data["Role"][i])
            break
        else:
            count = count + 1

    if ID_found == False:
        print("ID Not Found, Please Input Again")

print("~" + 'Welcome' + " " + name + "~")

print(datetime.now().strftime("Date: %d-%m-%y Time: %H:%M:%S"))
print("1. Start Shift")
print("2. Finish Shift")
print("3. Print Employee Times (Manager)")
print("4. Add Employee (Manager)")
print("5. Export Finished Employee Times (Manager)")
print("6. Exit Program")

option_found = False

while option_found == False:
    option = int(input("Choose an option (1-6):"))
    if 1 <= option <= 6:
        option_found = True
    else:
        print("Incorrect Entry, Try Again")

if option == 1:
    try:
        with open('EmployeeIN.csv') as file:
           reader = csv.reader(file)
           file.close()
    except IOError:
        with open('EmployeeIN.csv', 'ab') as csvf:
            writer = csv.writer(csvf)
            writer.writerow(["ID", "Date", "In", "Out", "Total"])
    with open('EmployeeIN.csv', 'rb') as f:
        reader = csv.reader(f)
        unique = True
        date = datetime.now().strftime("%d-%m-%y")

        for row in reader:
            if row[0] == ID and row[1] == date:
                unique = False
        f.close()

    with open('EmployeeIN.csv', 'ab') as f:
        writer = csv.writer(f)

        if unique == True:
            writer.writerow([ID,datetime.now().strftime("%d-%m-%y"),datetime.now().strftime("%H:%M")])
            print("Shift Started")
            
        if unique == False:
            print("Shift Already Started")
        f.close()

if option == 2:
    date = datetime.now().strftime("%d-%m-%y")
    time = datetime.now().strftime("%H:%M")
    t = "%H:%M"
    counter = 0
    found = False
    with open('EmployeeIN.csv', 'rb') as csvfile:
        reader = csv.reader(csvfile)
        employdata = []
        while found == False:
            for row in reader:
                if row[0] == ID and row[1] == date:
                    timein = row[2]
                    x = 0
                    found = True
                    for x in range(len(row)):
                       employdata.append(row[x])
            if found == False:
                print("Shift Not Started")
                sys.exit()
    
        #Add shift finishing time to list
        employdata.append(time)
        #Calculate the total time (Timefinal-Timeinitial) and convert into string
        timetotal = str(datetime.strptime(time,t)-datetime.strptime(timein,t))
        #Strip last 3 characters so it is formated 00:00
        timetotal = timetotal[:-3]
        #Add total time to list
        employdata.append(timetotal)
        
        
    with open('EmployeeOUT.csv', 'ab') as f:
        writer = csv.writer(f)
        writer.writerow(employdata)
        print("Shift Ended")
        f.close()

    with open('EmployeeIN.csv', 'rb') as inp, open('EmployeeIN2.csv', 'ab') as out:
        writer = csv.writer(out)
        for row in csv.reader(inp):
            if row[0] != ID:
                writer.writerow(row)
        inp.close()
        out.close()

    os.remove('EmployeeIN.csv')
    os.rename('EmployeeIN2.csv','EmployeeIN.csv')
    
if option == 3:
    if manager == 0:
        print("Permission Denied")
        
    if manager == 1:
        e_data = pd.read_csv("EmployeeOUT.csv")
        e_data.set_index("ID", inplace = True)
        print(e_data)

if option == 4:
    if manager == 0:
        print("Permission Denied")

    if manager == 1:
        number = 1
        unique = False
        first_name = raw_input("Input Employee First Name:")
        last_name = raw_input("Input Employee Last Name:")
        new_ID = str.upper("".join([last_name[:4],"00",str(number)]))
        with open('EmployeeList.csv', 'rb') as f:
            reader = csv.reader(f)
            while unique == False:
                for row in reader:
                    if row[0] == new_ID:
                        number = number + 1
                        new_ID = str.upper("".join([last_name[:4],"00",str(number)]))
                unique = True
        manager_status = raw_input("Manager? (Y/N):")
        role = 0
        if manager_status == "Y" or manager_status == "y":
            role = 1

        if manager_status == "N" or manager_status == "n":
            role = 0

        myCode = QR(data=new_ID, pixel_size=10)
        myCode.encode()
        filepath = myCode.filename
        QRfile = filepath.split("/")
        shutil.move(filepath,os.getcwd()+"/"+new_ID+".png")
        fromaddr = "frisbe05@live.com"

        e_mail = False
        while e_mail == False:
            toaddr = raw_input("Input Email Address (Case Sensitive):")
            answer = 0
            answer = raw_input('Are you sure you want to send the email to "'+toaddr+'"'+' (Y/N):')
            if answer == "Y" or answer == "y":
                e_mail = True
        
        msg = MIMEMultipart()
         
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = 'Employee QR Code'
         
        body = 'Attached is your QR code'
         
        msg.attach(MIMEText(body, 'plain'))
         
        filename = new_ID+".png"
        attachment = open(new_ID+".png", "rb")
        
         
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
         
        msg.attach(part)
         
        server = smtplib.SMTP('smtp.live.com', 587)
        server.starttls()
        server.login(fromaddr, "lala05")
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()

        with open('EmployeeList.csv','ab') as f:
            writer = csv.writer(f)
            writer.writerow([new_ID,(" ".join([first_name,last_name])),role])

if option == 5:
    if manager == 0:
        print("Permission Denied")

    if manager == 1:
        fromaddr = "friswk13@sbc.vic.edu.au"
        toaddr = "friswk13@sbc.vic.edu.au"
         
        msg = MIMEMultipart()
         
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = "Little Sister Employee Times"
         
        body = "Attached is this weeks employee's times."
         
        msg.attach(MIMEText(body, 'plain'))
         
        filename = "EmployeeOUT.csv"
        attachment = open("EmployeeOUT.csv", "rb")
         
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
         
        msg.attach(part)
         
        server = smtplib.SMTP('smtp.outlook365.com', 587)
        server.starttls()
        server.login(fromaddr, "jr9e9dtq")
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()

if option == 6:
    sys.exit()
