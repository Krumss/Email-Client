#!/usr/bin/python3

# Student name and No.:
# Development platform:
# Python version:

from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
import re
import os
import pathlib
import sys
import base64
import socket

#
# Global variables
#

# Replace this variable with your CS email address
YOUREMAIL = "u3568749@cs.hku.hk"
# Replace this variable with your student number
MARKER = '3035687493'

# The Email SMTP Server
SERVER = "testmail.cs.hku.hk"   #SMTP Email Server
SPORT = 25                      #SMTP listening port

# For storing the attachment file information
fileobj = None                  #For pointing to the opened file
filename = ''                   #For keeping the filename


#
# For the SMTP communication
#
def do_Send():
  # To be completed
  print('Send')

  mailserver = (SERVER, SPORT)
  clientSocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
  clientSocket.connect(mailserver)

  recv = clientSocket.recv(1024)
  print("Message after connection request:" +str(recv))
  if recv[:3] != '220':
    print('220 reply not received from server.')

  # Send HELO command and print server response.
  heloCommand = 'EHLO Alice\r\n'
  clientSocket.send(heloCommand.encode())
  recv1 = clientSocket.recv(1024)
  print(recv1)
  if recv1[:3] != '250':
    print('250 reply not received from server.')

  # Send MAIL FROM command and print server response.
  mailFrom = "MAIL FROM: <" + YOUREMAIL + "> \r\n"
  clientSocket.send(mailFrom.encode())
  recv2 = clientSocket.recv(1024)
  print("After MAIL FROM command: "+str(recv2))
  if recv2[:3] != '250':
    print('250 reply not received from server.')

  # Send RCPT TO command and print server response.
  rcptTo = "RCPT TO: <"+get_TO()+"> \r\n"
  clientSocket.send(rcptTo.encode())
  recv3 = clientSocket.recv(1024)
  print("After RCPT TO command: "+str(recv3))
  if recv3[:3] != '250':
    print('250 reply not received from server.')

  # Send DATA command and print server response.
  data = "DATA\r\n"
  clientSocket.send(data.encode())
  recv4 = clientSocket.recv(1024)
  print("After DATA command: "+str(recv4))
  if recv4[:3] != '354':
    print('354 reply not received from server.')

  f = "From: " + SERVER + "\r\n"
  clientSocket.send(f.encode())
  subject = "Subject: " + get_Subject() + "\r\n"
  clientSocket.send(subject.encode())
  to = "To: " + get_TO() + "\r\n"
  clientSocket.send(to.encode())
  if(get_CC != ""):
    cc = "Cc: " + get_CC() + "\r\n"
    clientSocket.send(cc.encode())
  clientSocket.send("\r\n".encode())
  msg = get_Msg() + "\r\n"
  clientSocket.send(msg.encode())
  end = ".\r\n"
  clientSocket.send(end.encode())
  recv5 = clientSocket.recv(1024)
  if recv5[:3] != '250':
    print('250 reply not received from server.')
  quitCommand = 'QUIT\r\n'
  clientSocket.send(quitCommand.encode())
  recv6 = clientSocket.recv(1024)
  if recv6[:3] != '221':
    print('221 reply not received from server.')
  clientSocket.close()


#
# Utility functions
#

#This set of functions is for getting the user's inputs
def get_TO():
  return tofield.get()

def get_CC():
  return ccfield.get()

def get_BCC():
  return bccfield.get()

def get_Subject():
  return subjfield.get()

def get_Msg():
  return SendMsg.get(1.0, END)

#This function checks whether the input is a valid email
def echeck(email):   
  regex = '^([A-Za-z0-9]+[.\-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+'  
  if(re.fullmatch(regex,email)):   
    return True  
  else:   
    return False

#This function displays an alert box with the provided message
def alertbox(msg):
  messagebox.showwarning(message=msg, icon='warning', title='Alert', parent=win)

#This function calls the file dialog for selecting the attachment file.
#If successful, it stores the opened file object to the global
#variable fileobj and the filename (without the path) to the global
#variable filename. It displays the filename below the Attach button.
def do_Select():
  global fileobj, filename
  if fileobj:
    fileobj.close()
  fileobj = None
  filename = ''
  filepath = filedialog.askopenfilename(parent=win)
  if (not filepath):
    return
  print(filepath)
  if sys.platform.startswith('win32'):
    filename = pathlib.PureWindowsPath(filepath).name
  else:
    filename = pathlib.PurePosixPath(filepath).name
  try:
    fileobj = open(filepath,'rb')
  except OSError as emsg:
    print('Error in open the file: %s' % str(emsg))
    fileobj = None
    filename = ''
  if (filename):
    showfile.set(filename)
  else:
    alertbox('Cannot open the selected file')

#################################################################################
#Do not make changes to the following code. They are for the UI                 #
#################################################################################

#
# Set up of Basic UI
#
win = Tk()
win.title("EmailApp")

#Special font settings
boldfont = font.Font(weight="bold")

#Frame for displaying connection parameters
frame1 = ttk.Frame(win, borderwidth=1)
frame1.grid(column=0,row=0,sticky="w")
ttk.Label(frame1, text="SERVER", padding="5" ).grid(column=0, row=0)
ttk.Label(frame1, text=SERVER, foreground="green", padding="5", font=boldfont).grid(column=1,row=0)
ttk.Label(frame1, text="PORT", padding="5" ).grid(column=2, row=0)
ttk.Label(frame1, text=str(SPORT), foreground="green", padding="5", font=boldfont).grid(column=3,row=0)

#Frame for From:, To:, CC:, Bcc:, Subject: fields
frame2 = ttk.Frame(win, borderwidth=0)
frame2.grid(column=0,row=2,padx=8,sticky="ew")
frame2.grid_columnconfigure(1,weight=1)
#From 
ttk.Label(frame2, text="From: ", padding='1', font=boldfont).grid(column=0,row=0,padx=5,pady=3,sticky="w")
fromfield = StringVar(value=YOUREMAIL)
ttk.Entry(frame2, textvariable=fromfield, state=DISABLED).grid(column=1,row=0,sticky="ew")
#To
ttk.Label(frame2, text="To: ", padding='1', font=boldfont).grid(column=0,row=1,padx=5,pady=3,sticky="w")
tofield = StringVar()
ttk.Entry(frame2, textvariable=tofield).grid(column=1,row=1,sticky="ew")
#Cc
ttk.Label(frame2, text="Cc: ", padding='1', font=boldfont).grid(column=0,row=2,padx=5,pady=3,sticky="w")
ccfield = StringVar()
ttk.Entry(frame2, textvariable=ccfield).grid(column=1,row=2,sticky="ew")
#Bcc
ttk.Label(frame2, text="Bcc: ", padding='1', font=boldfont).grid(column=0,row=3,padx=5,pady=3,sticky="w")
bccfield = StringVar()
ttk.Entry(frame2, textvariable=bccfield).grid(column=1,row=3,sticky="ew")
#Subject
ttk.Label(frame2, text="Subject: ", padding='1', font=boldfont).grid(column=0,row=4,padx=5,pady=3,sticky="w")
subjfield = StringVar()
ttk.Entry(frame2, textvariable=subjfield).grid(column=1,row=4,sticky="ew")

#frame for user to enter the outgoing message
frame3 = ttk.Frame(win, borderwidth=0)
frame3.grid(column=0,row=4,sticky="ew")
frame3.grid_columnconfigure(0,weight=1)
scrollbar = ttk.Scrollbar(frame3)
scrollbar.grid(column=1,row=1,sticky="ns")
ttk.Label(frame3, text="Message:", padding='1', font=boldfont).grid(column=0, row=0,padx=5,pady=3,sticky="w")
SendMsg = Text(frame3, height='10', padx=5, pady=5)
SendMsg.grid(column=0,row=1,padx=5,sticky="ew")
SendMsg.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=SendMsg.yview)

#frame for the button
frame4 = ttk.Frame(win,borderwidth=0)
frame4.grid(column=0,row=6,sticky="ew")
frame4.grid_columnconfigure(1,weight=1)
Sbutt = Button(frame4, width=5,relief=RAISED,text="SEND",command=do_Send).grid(column=0,row=0,pady=8,padx=5,sticky="w")
Atbutt = Button(frame4, width=5,relief=RAISED,text="Attach",command=do_Select).grid(column=1,row=0,pady=8,padx=10,sticky="e")
showfile = StringVar()
ttk.Label(frame4, textvariable=showfile).grid(column=1, row=1,padx=10,pady=3,sticky="e")

win.mainloop()
