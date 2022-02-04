
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import xlrd
import os
import PySimpleGUI as sg
sg.theme('Dark Blue 3')  


layout = [
    [sg.Text('Please enter your login credentials:', font='Default 14')],
    [sg.T('Username:', size=(15,1)), sg.Input(key='-USER-')],
    [sg.T('Password:', size=(15,1)), sg.Input(password_char='*', key='-PASSWORD-')],
    [sg.Button('Send'), sg.Cancel()]
]

window = sg.Window('Send Email', layout)
event, values = window.read()


str1 = os.path.dirname(__file__)
str2 = "Recruiters.xlsx"
loc = str1 + "/" + str2


wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


lst = []
for i in range(sheet.nrows):
    lst.append(sheet.cell_value(i, 2))



count = 0
for i in range(1, len(lst)):

    if(lst[i] != '' and count<=500):
        fromaddr = values['-USER-']

        toaddr = sheet.cell_value(i, 2)

       
        msg = MIMEMultipart()

        
        msg['From'] = fromaddr

  
        msg['To'] = toaddr


        msg['Subject'] = "Hi"

        body = "Test mail"

        msg.attach(MIMEText(body, 'plain'))

        str3 = "Resume.pdf"
        filename = str1 + "/" + str3
        attachment = open(filename, "rb")

        p = MIMEBase('application', 'octet-stream')

        p.set_payload((attachment).read())

        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(p)

        s = smtplib.SMTP('smtp.gmail.com', 587)

        s.starttls()

        s.login(fromaddr, values['-PASSWORD-'])
        sg.popup_quick_message('Sending your message... this will take a moment...', background_color='red')
        
        text = msg.as_string()

        s.sendmail(fromaddr, toaddr, text)

        s.quit()
        count = count +1
        print("Email to " + toaddr + " sent successfully")

print("Total Emails Sent: ")
print(count)
