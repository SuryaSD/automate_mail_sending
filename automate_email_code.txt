# Python code to illustrate Sending mail from  
# your Gmail account 
import pandas as pd
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText




userid="your email-id"
password="your password"
gmailid="@gmail.com"
outlookid="@outlook.com"

df = pd.read_csv (r'E:\project\bb\grades.csv')
email=(df.email[0:len(df)])
name=(df.name[0:len(df)])
roll=(df.Roll_No[0:len(df)])
grade=(df.grade[0:len(df)])

subject_for_PG="Welcome to PG"
subject_for_MSC="Welcome to MSC"
  
def gmail():
 import smtplib 
# creates SMTP session 
 s = smtplib.SMTP('smtp.gmail.com', 587) 
  
# start TLS for security 
 s.starttls() 
  
# Authentication 
 s.login(userid, password) 
  
# message to be sent
#newmsg=[]
#for i in range (len(df)):



#message = "Test"




 msg = MIMEMultipart()
 msg['Subject'] = subject_for_PG
#newmsg.append(message)
 msg1 = MIMEMultipart()
 msg1['Subject'] = subject_for_MSC

 filename = r"E:\Assignment\semester2\python\student.pdf"  # In same directory as script

 
 for i in range (len(df)):
  if ("PG" in roll[i]):
    #this is for adding subject

   TEXT = "hey"+' '+name[i]+" you are in PG"+" "+'\n' 
   
   msg.attach(MIMEText(TEXT, 'plain')) 
   #message= 'Subject: {}\n\n{}'.format(SUBJECT, TEXT)
   
     
# open the file to be sent  
   filename = "pandascheat.pdf"
   attachment = open(r"E:\Assignment\semester2\python\student.pdf", "rb") 
  
# instance of MIMEBase and named as p 
   p = MIMEBase('application', 'octet-stream') 
  
# To change the payload into encoded form 
   p.set_payload((attachment).read()) 
  
# encode into base64 
   encoders.encode_base64(p) 
   
   p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
  
# attach the instance 'p' to instance 'msg' 
   msg.attach(p) 
   text = msg.as_string() 
  

   s.sendmail("**",email[i],text) 

  elif ("MSC" in roll[i]):
        
    #this is for adding subject

   TEXT = "hey"+' '+name[i]+" you are in MSC"
   
   msg1.attach(MIMEText(TEXT, 'plain')) 
   #message= 'Subject: {}\n\n{}'.format(SUBJECT, TEXT)
   
     
# open the file to be sent  
   filename = "pandascheat.pdf"
   attachment = open(r"E:\Assignment\semester2\python\student.pdf", "rb") 
  
# instance of MIMEBase and named as p 
   p = MIMEBase('application', 'octet-stream') 
  
# To change the payload into encoded form 
   p.set_payload((attachment).read()) 
  
# encode into base64 
   encoders.encode_base64(p) 
   
   p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
  
# attach the instance 'p' to instance 'msg' 
   msg1.attach(p) 
   text = msg1.as_string() 
  
   s.sendmail("**",email[i],text) 
    
# terminating the session 
 s.quit() 
    
#gmail()


def outlook():
 import win32com.client as win32
 import psutil
 import os
 import subprocess
 import pandas as pd
 df = pd.read_csv (r'E:\oulook.csv')
 df
 i=len(df)
 i=len(df)
 email=(df.email[0:len(df)])
 name=(df.name[0:len(df)])
 #status=(df.status[0:len(df)])
 
# Drafting and sending email notification to senders. You can add other senders' email in the list
 def send_notification():
   for i in range (len(df)):
     outlook = win32.Dispatch('outlook.application')
     mail = outlook.CreateItem(0)
     mail.To = email[i]
     mail.Subject = subj
     mail.body = 'Welcome to Data Science'
     mail.send
     #open_outlook()
     
     
# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
     
 def open_outlook():
    try:
        subprocess.call(['"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"'])
        os.system("C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE");
    except:
        print("Outlook didn't open successfully")
  
    #send_notification()
 
# Checking if outlook is already opened. If not, open Outlook.exe and send email
 def kuchbhi():
  for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0
 
  if (flag == 1):
    send_notification()
  else:
    open_outlook()
    send_notification()

#send_notification()
 kuchbhi()
     
def decide():
    if gmailid in userid:
        gmail()
    else:
        outlook()

decide()