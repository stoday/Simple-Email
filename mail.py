import smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import json
import csv
from zipfile import ZipFile
import os


def get_sender_info():
    with open('sender_info.config') as read_file:
        sender_info = json.load(read_file)
        sender_email = sender_info['sender_email']
        password = sender_info['password']
    return sender_email, password

sender_email, password = get_sender_info()
print(sender_email)
print(password)


def get_receivers_info():
    with open('receivers_info.config', newline='') as read_file:
        receivers_info = csv.reader(read_file, delimiter=',')
        receiver_emails = [x[0] for x in receivers_info]
    return receiver_emails
receiver_mail = get_receivers_info()
print(receiver_mail[0])


def make_zip_file(zip_file_name, target_dictionary, filter_ext_name):
    with ZipFile(zip_file_name, 'w') as zip_obj:
        # Add multiple files to the zip
        files_list = os.listdir(target_dictionary)
        tar_files_list = [x for x in files_list if x.rfind(filter_ext_name)!=-1]
        print(tar_files_list)
        for tar_file in tar_files_list:
            zip_obj.write(tar_file)

make_zip_file('ttt.zip', '.', '.log')

# Create a multipart message
port = 587  # For SSL
smtp_server = 'smtp.office365.com'    # "smtp.gmail.com"
subject = "An email with attachment from Python-1"
body = "This is an email with attachment sent from Python"
# sender_email = "today@iii.org.tw"
sender_email = sender_email
# receiver_email = "stoday@gmail.com"
receiver_email = receiver_mail[0]
password = input("Type your password and press enter:")

message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
message["Bcc"] = receiver_email    # Recommended for mass emails
filename = "document.pdf"    # In same directory as script
# filename = "att_files.zip"    # In same directory as script

# Add body to email
message.attach(MIMEText(body, "plain"))

print('Attaching files...', end='')
with open(filename,'rb') as file:
    # Attach the file with filename to the email
    message.attach(MIMEApplication(file.read(), Name=filename))
    text = message.as_string()
print('done')    

# Create a secure SSL context
context = ssl.create_default_context()

# Try to log in to server and send email
try:
    print('Logging...', end='')
    server = smtplib.SMTP(smtp_server,port)
    server.ehlo() # Can be omitted
    server.starttls(context=context) # Secure the connection
    server.ehlo() # Can be omitted
    server.login(sender_email, password)
    print('done.')

    print('Sending...', end='')
    # Send email
    server.sendmail(sender_email, receiver_email, text)
    print('done')

except Exception as e:
    # Print any error messages to stdout
    print(e)

finally:
    server.quit()