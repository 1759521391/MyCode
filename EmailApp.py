#!/usr/bin/python3

# Student name and No.: zhhong 3035825784
# Development platform: Win10 64bit
# Python version:       python 3.10

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
YOUREMAIL = "zhhong@cs.hku.hk"
# Replace this variable with your student number
MARKER = '3035825784'

# The Email SMTP Server
SERVER = "testmail.cs.hku.hk"  # SMTP Email Server
SPORT = 25  # SMTP listening port

# For storing the attachment file information
fileobj = None  # For pointing to the opened file
filename = ''  # For keeping the filename


def split_by_comma(emails):
    '''
    split a string of emails
    '''
    lst = []
    for email in emails.split(','):
        email = email.strip()
        if email:
            lst.append(email)
    return lst


def read_line(sock):
    '''
    read line from socket, ending with \r\n
    '''
    line = ''
    while not line.endswith('\r\n'):
        bt = sock.recv(1)
        bt = str(bt, 'utf-8')
        line += bt
    print(line, end="")
    return line


def send_line(sock, line):
    '''
    send line by socket 
    '''
    print(line, end="")
    bts = line.encode('utf-8')
    ret = sock.send(bts)
    assert ret == len(bts)


def send_email(send_to, email_to, email_cc, email_bcc, email_subject, email_msg, file_name, file_byte):
    '''
    send email
    file_name: empty string for text type email; non-empty string for attachment type email.
    '''
    server_ip = socket.gethostbyname(SERVER)
    local_ip = '127.0.0.1'

    # (1) create socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    # (2) connect to the server
    sock.settimeout(10)
    try:
        sock.connect((server_ip, SPORT))
    except Exception as e:
        print(e)
        return "SMTP server is not availiable!"

    # (3) read connection message
    line = read_line(sock).strip()
    if not line.startswith('220'):
        # print(line)
        return f"Failed in connect!\n{line}"

    # (4) read info message
    send_line(sock, f"EHLO {local_ip}\r\n")
    line = read_line(sock).strip()
    while True:
        line = read_line(sock).strip()
        # print(line)
        if line.startswith('250 '):
            break

    # (5) send MAIL FROM
    send_line(sock, f"MAIL FROM:<{YOUREMAIL}>\r\n")
    line = read_line(sock).strip()
    if not line.startswith('250'):
        sock.close()
        return f"Fail in sending MAIL FROM\n{line}"

    # (6) send MAIL RCPT TO
    send_line(sock, f"RCPT TO:<{send_to}>\r\n")
    line = read_line(sock).strip()
    if not line.startswith('250'):
        sock.close()
        return f"Fail in sending RCPT TO\n{line}"

    # (7) send DATA
    send_line(sock, f"DATA\r\n")
    line = read_line(sock).strip()
    if not line.startswith('354'):
        sock.close()
        return f"Fail in sending DATA\n{line}"

    # (8) send header
    send_line(sock, f"From:{YOUREMAIL}\r\n")
    send_line(sock, f"To:{email_to}\r\n")
    send_line(sock, f"Subject:{email_subject}\r\n")
    if email_cc:
        send_line(sock, f"Cc:{email_cc}\r\n")
    if email_bcc:
        send_line(sock, f"Bcc:{email_bcc}\r\n")

    if file_name:
        ####################################
        # Send with Attachment
        ####################################
        send_line(sock, f"MIME-Version: 1.0\r\n")
        send_line(
            sock, f"Content-Type: multipart/mixed; boundary={MARKER}\r\n")

        # send message
        send_line(sock, f"\r\n")
        send_line(sock, f"--{MARKER}\r\n")
        send_line(sock, f"Content-Type: text/plain\r\n")
        send_line(sock, f"Content-Transfer-Encoding: 7bit\r\n")
        send_line(sock, f"\r\n")
        send_line(sock, f"{email_msg}\r\n")

        # send file
        send_line(sock, f"\r\n")
        send_line(sock, f"--{MARKER}\r\n")

        send_line(sock, f"Content-Type: application/octet-stream\r\n")
        send_line(sock, f"Content-Transfer-Encoding: base64\r\n")
        send_line(
            sock, f"Content-Disposition: attachment; filename={file_name}\r\n")
        send_line(sock, f"\r\n")
        send_line(sock, str(base64.encodebytes(file_byte), 'utf-8'))
        send_line(sock, f"--{MARKER}--\r\n")
    else:
        ####################################
        # Send without Attachment
        ####################################

        # add an empty line
        send_line(sock, f"\r\n")

        # (9) send CONTENT
        send_line(sock, f"{email_msg}\r\n")

    send_line(sock, f".\r\n")
    line = read_line(sock).strip()
    if not line.startswith('250'):
        sock.close()
        return f"Fail in sending period\n{line}"

    # send QUIT
    send_line(sock, f"QUIT\r\n")
    line = read_line(sock).strip()
    if not line.startswith('221'):
        sock.close()
        return f"Fail in sending quit\n{line}"
    sock.close()

    return "Successful"
#
# For the SMTP communication
#


def do_Send():
    '''
    repsonse to send click event
    '''

    # get data from UI
    email_to = get_TO().strip()
    email_cc = get_CC().strip()
    email_bcc = get_BCC().strip()
    email_subject = get_Subject().strip()
    email_msg = get_Msg().strip()

    if email_to == "":
        alertbox("Must enter the recipient's email")
        return
    if email_subject == "":
        alertbox("Must enter the subject")
        return
    if email_msg == "":
        alertbox("Must enter the message")
        return
    email_to_list = split_by_comma(email_to)
    email_cc_list = split_by_comma(email_cc)
    email_bcc_list = split_by_comma(email_bcc)

    # check email format
    for email in email_to_list:
        if not echeck(email):
            alertbox(f"Invalid To: Email - {email}")
            return
    for email in email_cc_list:
        if not echeck(email):
            alertbox(f"Invalid CC: Email - {email}")
            return
    for email in email_bcc_list:
        if not echeck(email):
            alertbox(f"Invalid BCC: Email - {email}")
            return
    # return
    bt = ""
    if fileobj:
        bt = fileobj.read()

    # send email to each recipent
    try:
        success = True
        for e in email_to_list + email_cc_list + email_bcc_list:
            msg = send_email(
                send_to=e,
                email_to=email_to,
                email_cc=email_cc,
                email_bcc=email_bcc,
                email_subject=email_subject,
                email_msg=email_msg,
                file_name=filename,
                file_byte=bt
            )
            alertbox(msg)
            if msg != "Successful":
                success = False
        if success:
            sys.exit(0)
    except Exception as e:
        # time out
        alertbox(f"SMTP server is not availiable!")

    return 0


#
# Utility functions
#

# This set of functions is for getting the user's inputs
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

# This function checks whether the input is a valid email


def echeck(email):
    regex = '^([A-Za-z0-9]+[.\-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+'
    if (re.fullmatch(regex, email)):
        return True
    else:
        return False

# This function displays an alert box with the provided message


def alertbox(msg):
    messagebox.showwarning(message=msg, icon='warning',
                           title='Alert', parent=win)

# This function calls the file dialog for selecting the attachment file.
# If successful, it stores the opened file object to the global
# variable fileobj and the filename (without the path) to the global
# variable filename. It displays the filename below the Attach button.


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
        fileobj = open(filepath, 'rb')
    except OSError as emsg:
        print('Error in open the file: %s' % str(emsg))
        fileobj = None
        filename = ''
    if (filename):
        showfile.set(filename)
    else:
        alertbox('Cannot open the selected file')

#################################################################################
# Do not make changes to the following code. They are for the UI                 #
#################################################################################


#
# Set up of Basic UI
#
win = Tk()
win.title("EmailApp")

# Special font settings
boldfont = font.Font(weight="bold")

# Frame for displaying connection parameters
frame1 = ttk.Frame(win, borderwidth=1)
frame1.grid(column=0, row=0, sticky="w")
ttk.Label(frame1, text="SERVER", padding="5").grid(column=0, row=0)
ttk.Label(frame1, text=SERVER, foreground="green",
          padding="5", font=boldfont).grid(column=1, row=0)
ttk.Label(frame1, text="PORT", padding="5").grid(column=2, row=0)
ttk.Label(frame1, text=str(SPORT), foreground="green",
          padding="5", font=boldfont).grid(column=3, row=0)

# Frame for From:, To:, CC:, Bcc:, Subject: fields
frame2 = ttk.Frame(win, borderwidth=0)
frame2.grid(column=0, row=2, padx=8, sticky="ew")
frame2.grid_columnconfigure(1, weight=1)
# From
ttk.Label(frame2, text="From: ", padding='1', font=boldfont).grid(
    column=0, row=0, padx=5, pady=3, sticky="w")
fromfield = StringVar(value=YOUREMAIL)
ttk.Entry(frame2, textvariable=fromfield, state=DISABLED).grid(
    column=1, row=0, sticky="ew")
# To
ttk.Label(frame2, text="To: ", padding='1', font=boldfont).grid(
    column=0, row=1, padx=5, pady=3, sticky="w")
tofield = StringVar()
ttk.Entry(frame2, textvariable=tofield).grid(column=1, row=1, sticky="ew")
# Cc
ttk.Label(frame2, text="Cc: ", padding='1', font=boldfont).grid(
    column=0, row=2, padx=5, pady=3, sticky="w")
ccfield = StringVar()
ttk.Entry(frame2, textvariable=ccfield).grid(column=1, row=2, sticky="ew")
# Bcc
ttk.Label(frame2, text="Bcc: ", padding='1', font=boldfont).grid(
    column=0, row=3, padx=5, pady=3, sticky="w")
bccfield = StringVar()
ttk.Entry(frame2, textvariable=bccfield).grid(column=1, row=3, sticky="ew")
# Subject
ttk.Label(frame2, text="Subject: ", padding='1', font=boldfont).grid(
    column=0, row=4, padx=5, pady=3, sticky="w")
subjfield = StringVar()
ttk.Entry(frame2, textvariable=subjfield).grid(column=1, row=4, sticky="ew")

# frame for user to enter the outgoing message
frame3 = ttk.Frame(win, borderwidth=0)
frame3.grid(column=0, row=4, sticky="ew")
frame3.grid_columnconfigure(0, weight=1)
scrollbar = ttk.Scrollbar(frame3)
scrollbar.grid(column=1, row=1, sticky="ns")
ttk.Label(frame3, text="Message:", padding='1', font=boldfont).grid(
    column=0, row=0, padx=5, pady=3, sticky="w")
SendMsg = Text(frame3, height='10', padx=5, pady=5)
SendMsg.grid(column=0, row=1, padx=5, sticky="ew")
SendMsg.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=SendMsg.yview)

# frame for the button
frame4 = ttk.Frame(win, borderwidth=0)
frame4.grid(column=0, row=6, sticky="ew")
frame4.grid_columnconfigure(1, weight=1)
Sbutt = Button(frame4, width=5, relief=RAISED, text="SEND", command=do_Send).grid(
    column=0, row=0, pady=8, padx=5, sticky="w")
Atbutt = Button(frame4, width=5, relief=RAISED, text="Attach",
                command=do_Select).grid(column=1, row=0, pady=8, padx=10, sticky="e")
showfile = StringVar()
ttk.Label(frame4, textvariable=showfile).grid(
    column=1, row=1, padx=10, pady=3, sticky="e")

win.mainloop()
