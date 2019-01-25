import datetime, smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email import encoders
from credentials import *

def send_results(file, email_file):
    #Sets up SMTP settings and creates mailing list
    email_list=open(email_file, "r")
    email_addresses = email_list.read().split(';')

    #Create container for message and attach file
    msgRoot=MIMEMultipart('related')
    msgRoot['From']='mooneyryanj@gmail.com'
    msgRoot['To']=", ".join(email_addresses)
    msgRoot['Subject']='Asset Location Results for '+str(datetime.date.today())
    message = 'Attached please find the asset location results for '+str(datetime.date.today()+'</br></br>Signed,</br>The Pump Whisperer</br></br>')
    msgText=MIMEText(message, 'html')
    msgRoot.attach(msgText)

    attachment=open(file, 'rb')
    xlsx=MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    xlsx.set_payload(attachment.read())
    encoders.encode_base64(xlsx)
    xlsx.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
    msgRoot.attach(xlsx)

    #Send emails
    smtp = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
    smtp.starttls()
    smtp.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
    smtp.sendmail(EMAIL_HOST_USER, email_addresses, msgRoot.as_string())
    smtp.quit()

