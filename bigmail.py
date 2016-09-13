#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Configuration part ###########################################################

# Do not use non ASCII characters...
subject = 'Python-Tools'

toaddr = 'christian.dreier@csa-germany.de'
fromaddr = 'flipper.x3r@googlemail.com'
host = 'smtp.googlemail.com'
failed_at = 0
################################################################################

import sys
import os
import os.path

import tkinter
import tkinter.simpledialog

import smtplib
import mimetypes

import email
import email.mime.base
import email.mime.multipart
import email.mime.text
import email.mime.image
import email.mime.audio

att_dir = 'attachment'
msg_file = 'msg.txt'

# Helper functions #############################################################
def error(msg):
    print(msg + '\n', file = sys.stderr)
    exit(1)
    
def file2attachment(path):
    # Get file type from file name
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        # No guess could be made, or the file is encoded (compressed), so
        # use a generic bag-of-bits type.
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    
    # Open the file in an appropriate way
    if maintype == 'text':
        fp = open(path)
        # Note: we should handle calculating the charset
        att = email.mime.text.MIMEText(fp.read(), _subtype = subtype)
        fp.close()
    elif maintype == 'image':
        fp = open(path, 'rb')
        att = email.mime.image.MIMEImage(fp.read(), _subtype = subtype)
        fp.close()
    elif maintype == 'audio':
        fp = open(path, 'rb')
        att = email.mime.audio.MIMEAudio(fp.read(), _subtype = subtype)
        fp.close()
    else:
        fp = open(path, 'rb')
        att = email.mime.base.MIMEBase(maintype, subtype)
        att.set_payload(fp.read())
        fp.close()
        
        # Encode the payload using Base64
        email.encoders.encode_base64(att)
        
    # Set the filename parameter
    fname = os.path.basename(path)
    att.add_header('Content-Disposition', 'attachment', filename = fname)
    
    return att

def getpasswd():
    '''Opens a graphical dialog to type in a password.
    
    Returns:
        A string, containing the typed password.
    '''
    root = tkinter.Tk()
    root.withdraw()
    passwd = tkinter.simpledialog.askstring(
        'Enter Password',
        'Password: ',
        show = '*'
    )
    root.destroy()
    return passwd
################################################################################

# Main part ####################################################################
with open(msg_file, 'r') as f:
    body = f.read()
#print(body)

dirlist = os.listdir(att_dir)
dirsize = len(dirlist)
if dirsize == 0:
    error('No files to send.')
    
passwd = getpasswd()

for i, file in enumerate(dirlist):
    if i < (failed_at - 1):
        continue
    print('Try to send mail {} of {}: {}'.format(i + 1, dirsize, file))
    
    path = os.path.join(att_dir, file)
    
    # Build the e-mail
    mail = email.mime.multipart.MIMEMultipart()
    mail['Subject'] = subject + ' {}/{}'.format(i + 1, dirsize)
    mail['To'] = toaddr
    mail['From'] = fromaddr
    mail.attach(email.mime.text.MIMEText(body, _charset = 'utf8'))
    mail.attach(file2attachment(path))
    
    # Send the e-mail ######################################################
    
    # Connect to server
    try:
        server = smtplib.SMTP(host = host)
    except smtplib.SMTPConnectError as exc:
        error('Could not connect to server.\n' + str(exc))
        
    #server.set_debuglevel(1)
    
    try:
        # Start encryption
        try:
            server.starttls()
        except smtplib.SMTPHeloError as exc:
            error(
                'The server didn’t reply properly to the HELO greeting.\n'
                + str(exc)
            )
        except smtplib.SMTPException as exc:
            error(
                'The server does not support the STARTTLS extension.\n'
                + str(exc)
            )
        except RuntimeError as exc:
            error(
                'SSL/TLS support is not available to your Python interpreter.\n'
                + str(exc)
            )
        
        # Login
        try:
            server.login(fromaddr, passwd)
        except smtplib.SMTPHeloError as exc:
            error(
                'The server didn’t reply properly to the HELO greeting.\n'
                + str(exc)
            )
        except smtplib.SMTPAuthenticationError as exc:
            error(
                'The server didn’t accept the username/password combination.\n'
                + str(exc)
            )
        except smtplib.SMTPException as exc:
            error('No suitable authentication method was found.\n' + str(exc))
            
        # Do the actual sending
        try:
            err = server.send_message(mail)
        except smtplib.SMTPRecipientsRefused as exc:
            error('All recipients were refused.\n' + str(exc))
        except smtplib.SMTPHeloError as exc:
            error(
                'The server didn’t reply properly to the HELO greeting.\n'
                + str(exc)
            )
        except smtplib.SMTPSenderRefused as exc:
            error('The server didn’t accept the sender address.\n' + str(exc))
        except smtplib.SMTPDataError as exc:
            error(
                'The server replied with an unexpected error code.\n'
                + str(exc)
            )
            
        if len(err) > 0:
            for key, val in err.items():
                print(
                    'Sending to {} failed with following error: {}'.format(
                        key,
                        val
                    ),
                    file = sys.stderr
                )
    finally:
        server.quit()
    ########################################################################
################################################################################
