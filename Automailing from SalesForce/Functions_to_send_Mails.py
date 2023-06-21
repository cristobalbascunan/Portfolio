#!/usr/bin/env python
# coding: utf-8

import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import COMMASPACE, formatdate
from email import encoders


def send_mail(send_from: str, send_to: list, subject: str, message: str, signature: str, files: list,
              username:str, password:str, server="localhost", port=587, use_tls=True) -> None:
    """Compose and send email with provided info and attachments.
    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body as string written in html
        signature (str): signature as string written in html
        files (list[str]): list of file paths to be attached to email
        username (str): server auth username
        password (str): server auth password
        server (str): mail server host name
        port (int): port number
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    #Set sender
    msg['From'] = send_from
    #Set receiver
    msg['To'] = ", ".join(send_to)
    #Set formatdate
    msg['Date'] = formatdate(localtime=True)
    #Set subject
    msg['Subject'] = subject
    #Set message and signature. In order to apply format, that message has to be formatted in HTML.
    full_message = str(message+" \n"+signature)
    #Add to the mail the content of the message.
    msg.attach(MIMEText(full_message, 'html', 'utf-8'))
    #If there's need to add attachments, there will procesed in this iteration
    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msg.attach(part)
    #Set server and port
    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    #Set Username (email) and password
    smtp.login(username, password)
    #Send mail
    smtp.sendmail(send_from, send_to, msg.as_string())
    #Quit connection
    smtp.quit()