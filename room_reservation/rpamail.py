import pymsgbox
import pandas as pd         
import sys
import pyodbc
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
import time
import ntpath
import numpy as np
import xlsxwriter

pathfile = r'f:\bot\nutriex\rpabooking\pathrobo.txt'
pathway = (open(pathfile, 'r').read())
emailTo = []
emailCc = []
emailTo.append(open(pathway + r'emailTo.txt', 'r').read())
emailCc.append(open(pathway + r'emailCc.txt', 'r').read())
subject = (open(pathway + r'subject.txt', 'r').read())
message = (open(pathway + r'message.txt', 'r').read())
arquivo = (pathway + r'code.txt')


msg = MIMEMultipart()
msg['Subject'] = subject
msg['From'] = u'Autoatendimento Reserva de Salas <robot@nutriex.com.br>'
msg['To'] = ', '.join(emailTo)
msg['Cc'] = ', '.join(emailCc)
body = message
msg.attach(MIMEText(body, 'plain'))
filename = ntpath.basename(arquivo)
attachment = open(arquivo, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(part)
server = smtplib.SMTP('correio.mileniofarma.com.br', 587)
server.starttls()
server.login('robot@nutriex.com.br', '111111')
text = msg.as_string()
server.sendmail('robot@nutriex.com.br', emailTo + emailCc, text)
server.quit() 

sys.exit(0)