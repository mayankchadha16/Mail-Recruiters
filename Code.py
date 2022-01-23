# Python code to illustrate Sending mail with attachments
# from your Gmail account

# libraries to be imported.....smtplib, xlrd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import xlrd
import os

# Give the location of the file
str1 = os.path.dirname(__file__)
str2 = "Recruiters.xlsx"
loc = str1 + "/" + str2

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Storing emails in a list
lst = []
for i in range(sheet.nrows):
    lst.append(sheet.cell_value(i, 2))

#print(lst)

count = 0
for i in range(1, len(lst)):

    if(lst[i] != '' and count<=500):
        fromaddr = "anshuliiitb16@gmail.com"

        toaddr = sheet.cell_value(i, 2)

        # instance of MIMEMultipart
        msg = MIMEMultipart()

        # storing the senders email address
        msg['From'] = fromaddr

        # storing the receivers email address
        msg['To'] = toaddr

        # storing the subject
        msg['Subject'] = "Hi"

        # string to store the body of the mail
        body = "Test mail"

        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # open the file to be sent
        str3 = "Resume.pdf"
        filename = str1 + "/" + str3
        attachment = open(filename, "rb")

        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

        # To change the payload into encoded form
        p.set_payload((attachment).read())

        # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # attach the instance 'p' to instance 'msg'
        msg.attach(p)

        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # start TLS for security
        s.starttls()

        # Authentication
        s.login(fromaddr, "anshuldon12")

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # sending the mail
        s.sendmail(fromaddr, toaddr, text)

        # terminating the session
        s.quit()
        count = count +1
        print("Email to " + toaddr + " sent successfully")

print("Total Emails Sent: ")
print(count)
