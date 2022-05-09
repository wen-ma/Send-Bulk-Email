import os
import sys
import smtplib
import textwrap
from email import encoders
from email.header import Header  # Set up the Email header and subject
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename

import pandas as pd
import logging


def get_logger(filename, verbosity=1, name=None):
    """Define logger configuration to display on the console and log to file"""
    level_dict = {0: logging.DEBUG, 1: logging.INFO, 2: logging.WARNING}
    formatter = logging.Formatter("[%(asctime)s][%(levelname)s] %(message)s")
    logger = logging.getLogger(name)
    logger.setLevel(level_dict[verbosity])
    fh = logging.FileHandler(filename, "a+")
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    sh = logging.StreamHandler()
    sh.setFormatter(formatter)
    logger.addHandler(sh)
    return logger


log_file = get_logger(filename=f"log")


def add_attach(attach_file):
    """Setup to attach the file and keep the base name of the attached file"""
    att = MIMEBase('application', 'octet-stream')
    att.set_payload(open(attach_file, 'rb').read())
    # att.add_header('Content-Disposition', 'attachment', filename=('gbk', '', attach_file.split("\\")[-1]))
    att.add_header('Content-Disposition', 'attachment; filename="%s"' % basename(attach_file))
    encoders.encode_base64(att)
    return att


def addresses(addrstring):
    """Split in semicolon, strip surrounding whitespace."""
    return [x.strip() for x in addrstring.split(';')]


# Set up smtp server
smtpServer = 'smtp.gmail.com'
# smtpServer = 'mail.bluetie.com'
# Username and password of the sender
username = input('Please enter your username:')
password = input('Please enter your password:')
current_path = os.getcwd()
df = pd.read_excel('{0}/Email.xlsx'.format(current_path))
df = df.fillna(value='')
df_col = df.values.tolist()
try:
    session = smtplib.SMTP_SSL(smtpServer, 465)  # use gmail with port
    # session = smtplib.SMTP('213.182.20.68')
    session.ehlo()
    # session.starttls()  # enable security
    session.login(username, password)  # login with mail_id and password
    for df_col in df_col:
        message = MIMEMultipart()
        message.attach(MIMEText(textwrap.dedent(df_col[3]).strip(), 'html', 'utf-8'))  # message body
        message['From'] = username  # sender
        message['To'] = df_col[0]  # recipients
        message['cc'] = df_col[1]  # cc recipients
        message['Subject'] = Header(df_col[2], 'utf-8')  # subject
        if os.path.exists(df_col[4]):
            message.attach(add_attach(df_col[4]))  # add attachments into message container
            # Assume the signature image is in the current directory
            fp = open(r'signature.PNG', 'rb')
            msgImage = MIMEImage(fp.read())
            fp.close()
            # Define the image's ID as referenced above
            msgImage.add_header('Content-ID', '<image1>')
            message.attach(msgImage)
        else:
            log_file.info(f"Mail could not be sent to {df_col[0]} successfully without attachment!")
            continue
        session.sendmail(username, addresses(message['To']) + addresses(message['cc']), message.as_string())
        # session.sendmail(username, message['cc'], message.as_string())
        log_file.info(f"Mail sent successfully!")
    session.quit()
except smtplib.SMTPException:
    log_file.info("Something went wrong", exc_info=sys.exc_info())
    # log_file.debug("Something went wrong", exc_info=True)
