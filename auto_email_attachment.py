import os
import smtplib
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email.MIMEMultipart import MIMEMultipart
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def send(name, passwd, subject, body, addrto, q):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = name + '@gmail.com'
    msg['To'] = addrto

    text = MIMEText(body)
    msg.attach(text)

    if(q == "y"):
        img = askopenfilename()
        img_data = open(img, 'rb').read()
        image = MIMEImage(img_data, name=os.path.basename(img))
        msg.attach(image)

    print("connecting")
    s = smtplib.SMTP('smtp.gmail.com', '587')
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(name, passwd)
    print("connected")
    s.sendemail(name + '@gmail.com', addrto, msg.as_string())
    print("sent")
    s.quit()

Tk().withdraw()
name = input("your gmail> ")
passwd = input("password of gmail> ")
subject = input("subject> ")
body = input("text> ")
q = input("do you want to attach a file y/n> ")
addrto = input("reciepent email?> ")

send(name, passwd, subject, body, addrto, q)










