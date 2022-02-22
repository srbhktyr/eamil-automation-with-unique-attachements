import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

df = pd.read_excel('excel_or_csv_file_containig_email_ids_name_etc'.xlsx) # use csv or excel file to enter the sender details

email_list = df['Email Address'] # load senders email ids into a list
name = df["Name"]

i = 0
for x in email_list:
    y = name[i]
    msg = MIMEMultipart()

    from_add = "sender's email id"
    to_add = x
    subject = 'Email_subject'
    content = "Dear {},\n\n " \
              " Thanks and Regards: \n Mr. XYZ Singh\n Director, Your organization".format(y)
    body = MIMEText(content, 'plain')
    msg.attach(body)
    msg.add_header('subject', subject)
    #########

    #rename the file name with common text and add numbering after the name

    # for ex: file-1, file-2 etc

    ##########
    file_name = 'file-' + str((i + 1)) + '.pdf' #file name name formating.
    i += 1

    with open(file_name, 'rb') as f:
        attachment = MIMEApplication(f.read(), Name=basename(file_name))
        attachment['Content-Disposition'] = 'attchment; file_name="{}"'.format(basename(file_name))
        msg.attach(attachment)

    server = smtplib.SMTP('smtp.outlook.com', 587)
    server.starttls()

    server.login('sender email id', 'password')

    server.send_message(msg, from_addr=from_add, to_addrs=[to_add])
