import datetime
import ntpath
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from pytz import timezone

#etapa de transformacao dos dados (separador de colunas, milhar e decimal)
pathcoleta = u'f:\\bot\\rpa005\\'
files = pathcoleta + u'rpa005.csv'
time = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')

df = pd.read_csv(files, skiprows=[1], sep=',', header=0, error_bad_lines=False, dtype='unicode')
cols = df.columns[[4,5,8,9,10,11,12]]
df[cols] = df[cols].apply(pd.to_numeric, errors='coerce', axis=0)
df.to_csv(files, sep=';', encoding='utf-8-sig', quotechar='"', decimal=',', index=False)

# df = pd.read_csv(files, skiprows=[1], sep=',', header=0, error_bad_lines=False, dtype='unicode')
# df.to_csv(files, sep=',', encoding='utf-8', quotechar='"', decimal='.', index=False)

now = datetime.datetime.now()
arquivo = pathcoleta + 'rpa005_' + now.strftime('%d%m%y') + '.csv'
os.rename(files, arquivo)

#ENVIA EMAIL DO RELATORIO FINAL PARA OS DESTINATARIOS
email = 'robot@nutriex.com.br'
send_to_email = ['cristina.guimaraes@nutriex.com.br']
cc = ['rodrigo.matos@nutriex.com.br','luiz.xavier@nutriex.com.br']
subject = 'Relatorio Diario de Ordens de Producao Encerradas desde Out/2019 | ' + str(time)
message = 'Segue o relatorio extraido do SAP em formato CSV. Equipe Inteligencia 2020.'
password = '111111'

msg = MIMEMultipart()

msg['Subject'] = subject
msg['From'] = u'Robo Inteligencia <robot@nutriex.com.br>'
msg['To'] = ', '.join(send_to_email)
msg['Cc'] = ', '.join(cc)

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
server.sendmail(email, send_to_email+cc, text)
server.quit()
