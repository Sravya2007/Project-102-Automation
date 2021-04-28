#AUTOMATING OUTLOOK OFFICE 365 EMAILS
#This program sends an email from and to the specified addresses with a randomly selected message from a text file at a certain time everyday
#Works only with Microsoft Office365 Outlook

#SMTP (Simple Mail Transfer Protocol) Library
import smtplib
#required email sending modules of MIME (Multipurpose Internet Mail Extensions)
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
#os module
import os.path
#module to get passwords securely
import getpass
#random and time modules
import random
import time
#install schedule module: pip install schedule or python3 -m pip install schedule
import schedule

#getting the email address and password of the user
#email address must be an Outlook address
email_sender = input("Enter your Outlook email address: ")
credential = getpass.getpass("Enter the password to your email: ")

def send_email(email_recipient, email_subject, email_message):
    #email object with multiple parts like from, to, subject, message etc.
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    #attaching message provided as parameter to the email
    msg.attach(MIMEText(email_message, 'plain'))

    #setting up an Office365 SMTP server with port number 587 which is used to send emails
    server = smtplib.SMTP('smtp.office365.com', 587)
    #initialzing server
    server.ehlo()
    server.starttls()
    #logging into the Outlook server with details provided by user
    server.login(email_sender, credential)
    #return the entire message as a string
    text = msg.as_string()
    #sending the mail to recipient
    server.sendmail(email_sender, email_recipient, text)
    print('Email sent!')
    #terminate the SMTP session
    server.quit()
    return True

#picks up a random line from a text file (quotes in this case)
def random_line(file):
    lines = open(file).read().splitlines()
    #choose a random element from a non-empty sequence
    return random.choice(lines)
   
msg = random_line('messages.txt')
#getting recipient email address from user
recipient = input("Enter the recipient's email address: ")
#schedule an email at 10:00 AM to the recipient with a random quote
schedule.every().day.at("10:00").do(send_email, recipient, 'Inspirational message from ' + email_sender, "Dear " + recipient + " , \n\n" + "Good morning! 🌞 \n\n I have a motivating message just for you: \n\n" + msg)

#cause the schedule to wait till the desired time
while True:
    schedule.run_pending()
    #delays execution by 1 second
    time.sleep(1)