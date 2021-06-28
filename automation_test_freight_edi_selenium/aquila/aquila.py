from selenium.webdriver.common.keys import Keys
from selenium import webdriver
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

pathcoleta = 'F:\\python\\fretes\\coleta\\'
now = datetime.datetime.now()

files = glob.glob(pathcoleta + '*.sswweb')
for f in files:
    os.remove(f)

files = glob.glob(pathcoleta + '*.csv')
for f in files:
    os.remove(f)

driver = webdriver.Chrome('F:\\python\\fretes\\aquila\\chromedriver.exe')
driver.get("http://sistema.ssw.inf.br/bin/ssw0422")

# ACESSO A PAGINA DE LOGIN
u = driver.find_element_by_name('f1')
u.send_keys('AQL')
p = driver.find_element_by_name('f2')
p.send_keys('03446420118')
q = driver.find_element_by_name('f3')
q.send_keys('deivid')
t = driver.find_element_by_name('f4')
t.send_keys('deiv1010')
v = driver.find_element_by_id('5')
v.click()

# ACESSO DO MODELO DE RELATORIO 102
time.sleep(2)
y = driver.find_element_by_id('3')
y.send_keys('102')
time.sleep(2)

# ACESSO AO FILTRO DOS CTEs POR CPF
driver.switch_to.window(driver.window_handles[1])
j = driver.find_element_by_id('3')
j.send_keys('06219757000157')
time.sleep(2)

# ACESSO AO FILTRO DOS CTEs POR DATA/FORMATO EXPORTACAO
driver.switch_to.window(driver.window_handles[2])
k = driver.find_element_by_id('18')
k.send_keys(now.strftime('01%m%y'))
d = driver.find_element_by_id('22')
d.send_keys(Keys.BACKSPACE)
d.send_keys('E')
m = driver.find_element_by_id('24')
m.click()

#ENCERRA CAIXA DE DIALOGO DE CONCLUSAO RELATORIO
time.sleep(2)
n = driver.find_element_by_id('0')
n.click()

#AGUARDA UM TEMPO PARA EXTRACAO E FECHA TUDO
time.sleep(2)
driver.quit()

files = glob.glob(pathcoleta + '*.sswweb')
for file in files:
    os.rename(file, pathcoleta + 'aquila_'+now.strftime('%d%m%y')+'.csv')
    break

email = 'rodrigo.matos@mw.far.br'
send_to_email = ['andrea.ferraz@mw.far.br']
cc = ['rodrigo.matos@mw.far.br','rafael.augusto@vdmoplog.com.br']
subject = 'Relatorio Diario de Fretes (Transportadora AQUILA)'
message = 'Ola! Segue o relatorio extraido do portal da transportadora em formato CSV. Att., Rodrigo Matos.'
file_location = pathcoleta + 'aquila_'+now.strftime('%d%m%y')+'.csv'
password = 'juizdefora'

msg = MIMEMultipart()

msg['Subject'] = subject
msg['From'] = email
msg['To'] = ', '.join(send_to_email)
msg['Cc'] = ', '.join(cc)

body = message

msg.attach(MIMEText(body, 'plain'))

filename = ntpath.basename(file_location)
attachment = open(file_location, "rb")

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