from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import smtplib
from email.header import Header
from email import encoders
from datetime import datetime
import os


#step 1: backup excel file, save file to backup folder
a= os.scandir('C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA')
b= next(a)
c= next(a)
d= next(a)
print (d)
fileName= str(d)

strfileName=fileName[fileName.index('M'): fileName.index('>')-1]

strfileNamePDF= strfileName[0: strfileName.index('.')]+'.pdf'

print (strfileName)
file_to_copy='C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\'+strfileName
dest_dir= 'C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\Backup'
arc_dir= 'C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\Archive\\'
fileNamePDFDir=arc_dir+strfileNamePDF

now= datetime.now()
#date_time=now.strftime("%Y%m%d%H:%M:%S")
date_time=now.strftime("%Y%m%d_%H%M%S")

strfileName1= date_time+'_'+strfileName
strfileName1Dir=arc_dir+strfileName1

# step 10 send email with the attahced file(orignal excel file, updated excel file, PDF file)


sender = 'xuwei.blue@gmail.com'
receivers = ['xuwei.blue@gmail.com']

# content
msg = MIMEMultipart()
# msg = MIMEText('this is for python test', 'plain','utf-8')
msg['from'] = Header("test", 'utf-8')
msg['to'] = Header("test2", 'utf-8')
subject = 'python test'
msg['subject'] = Header(subject, 'utf-8')

msg.attach(MIMEText('this is for python test', 'plain', 'utf-8'))

with open(fileNamePDFDir, "rb") as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={strfileNamePDF}",
    )
    msg.attach(part)

with open(file_to_copy, "rb") as attachment:
    part1 = MIMEBase('application', 'octet-stream')
    part1.set_payload(attachment.read())

    encoders.encode_base64(part1)
    part1.add_header(
        "Content-Disposition",
        f"attachment; filename={strfileName}",
    )
    msg.attach(part1)

with open(strfileName1Dir, "rb") as attachment:
    part2 = MIMEBase('application', 'octet-stream')
    part2.set_payload(attachment.read())

    encoders.encode_base64(part2)
    part2.add_header(
        "Content-Disposition",
        f"attachment; filename={strfileName1}",
    )
    msg.attach(part2)

try:

    server = smtplib.SMTP()

    server.connect('servername', 25)

    server.sendmail(sender, receivers, msg.as_string())
    print("sucess, step 10 done")
    server.quit()
except smtplib.SMTPException:
    # except SMTPResponseException as e:
    # print (e.smtp_code)
    # print (e.smtp_error)
    print("error for sending")