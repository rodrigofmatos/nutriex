from pytz import timezone
import time
import os
import glob
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
import datetime
import pandas as pd

#etapa de transformacao dos dados (separador de colunas, milhar e decimal)
pathcoleta = u'f:\\bot\\rpa001\\'
files = pathcoleta + u'rpa001.csv'
time = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')

df = pd.read_csv(files, skiprows=[1], sep=';', header=0, error_bad_lines=False, dtype='unicode')
cols = df.columns[[8,9,10]]
df[cols] = df[cols].apply(pd.to_numeric, errors='coerce', axis=0)
df.to_csv(files, sep=';', encoding='utf-8-sig', quotechar='"', decimal=',', index=False)

now = datetime.datetime.now()
arquivo = pathcoleta + 'rpa001_' + now.strftime('%d%m%y') + '.csv'
os.rename(files, arquivo)

#ENVIA EMAIL DO RELATORIO FINAL PARA OS DESTINATARIOS
email = 'robot@nutriex.com.br'
send_to_email = ['rodrigo.matos@nutriex.com.br'] #['luiz.xavier@mw.far.br']
cc = ['rodrigo.matos@nutriex.com.br']
bcc = ['rodrigo.matos@nutriex.com.br'] #['katila.souza@nutriex.com.br']
subject = 'Relatorio Diario de Estoque | ' + str(time)
message = 'Segue o estoque extraido do SAP em formato CSV. Att., Equipe Inteligencia 2020.'
password = '111111'

msg = MIMEMultipart()

msg['Subject'] = subject
msg['From'] = u'Robo Inteligencia <robot@nutriex.com.br>'
msg['To'] = ', '.join(send_to_email)
msg['Cc'] = ', '.join(cc)
msg['Bcc'] = ', '.join(bcc)

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
server.login(email, password)
text = msg.as_string()
server.sendmail(email, send_to_email+cc+bcc, text)
server.quit()
