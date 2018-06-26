import pandas as pd
from selenium import webdriver
import time
import glob
import os
import smtplib
import os.path as op
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

# Identify Variables
downloads_folder = 'INSERT FOLDER PATH WHERE DOC IS LOCATED'
doc = 'INSERT ENTIRE PATH TO DOC.XLSX'
browser = webdriver.Chrome(executable_path='PATH TO CHROMEDRIVER.EXE')
sheet1 = business_id1 = 'SHEET 1 NAME'
sheet2 = business_id2 = 'SHEET 2 NAME'
fb_url = 'https://business.facebook.com/login'
fb_username = 'FACEBOOK USERNAME'
fb_password = 'FACEBOOK PASSWORD'

# Create Dataframe of each sheet
df1 = pd.read_excel(io=doc, sheet_name=sheet1, header=0)  # Assumes no header rows
df2 = pd.read_excel(io=doc, sheet_name=sheet2, header=0)

# Create list of hyperlinks from first column of each Dataframe
links1 = df1.values.T[0].tolist()  # Assumes hyperlinks in first data column
links2 = df2.values.T[0].tolist()

# Login to Facebook
browser.get(fb_url)
browser.find_element_by_id("email").send_keys(fb_username)
browser.find_element_by_id("pass").send_keys(fb_password)
browser.find_element_by_id("loginbutton").click()

# Loop through hyperlinks
for url in links1:
    urls = "'" + url + "'"
    print(urls, '\n')
    browser.get(url)
    time.sleep(2)

for url in links2:
    urls = "'" + url + "'"
    print(urls, '\n')
    browser.get(url)
    time.sleep(2)

browser.close()

# Define Send Email Function


def send_mail(send_from, send_to, subject, message, files=[],
              server="", port=587, username='', password=''):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (str): to name
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        for pathz in path:  # Loop through files per path given by FILE IDENTIFIERS below in send_email command
            part = MIMEBase('application', "octet-stream")
            with open(pathz, 'rb') as fil:
                part.set_payload(fil.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            'attachment; filename="{}"'.format(op.basename(pathz)))
            msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    smtp.ehlo()  # Must say ehlo() to gmail
    smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    print ('email sent')  # Should see this on console if successful
    smtp.quit()

# Send Email Command


send_mail(send_from='INSERT SEND_FROM EMAIL ADDRESS',
          subject="INSERT SUBJECT",
          message="INSERT MESSAGE",
          send_to='INSERT SEND_TO EMAIL ADDRESS',
          files=[glob.glob(os.path.join(downloads_folder, '*INSERT FILE IDENTIFIER 1*')),
                 glob.glob(os.path.join(downloads_folder, '*INSERT FILE IDENTIFIER 2*'))], server='smtp.gmail.com',
          port=587, username='GMAIL USERNAME', password='GMAIL PASSWORD')
