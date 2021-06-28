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

now = datetime.datetime.now()
pathcoleta = 'F:\\python\\fretes\\coleta\\'

files = glob.glob(pathcoleta + '*.xls')
for f in files:
    os.remove(f)

files = glob.glob(pathcoleta + 'tnt*.*')
for f in files:
    os.remove(f)

driver = webdriver.Chrome('F:\\python\\fretes\\tnt\\chromedriver.exe')
driver.get("http://radar.tntbrasil.com.br/radar/public/login")

# ACESSO A PAGINA DE LOGIN
u = driver.find_element_by_id("login")
u.send_keys('luiz.fernando@mw.far.br')
p = driver.find_element_by_id("senha")
p.send_keys('milenio00')
w = driver.find_element_by_css_selector("a[id='login']").click()
time.sleep(2)

continua = True
while continua:
    try:
        driver.find_element_by_xpath("//*[@id='menuPrivado']/li[4]/a[2]/img").get_attribute()
    except:
        pass
        if driver.find_element_by_xpath("//*[@id='menuPrivado']/li[4]/a[2]/img").is_displayed(): break
        time.sleep(0.1)

driver.find_element_by_xpath("//*[@id='menuPrivado']/li[4]/a[2]/img").click()

# PARAMETRIZACAO DOS FILTROS
driver.find_element_by_id("strPeriodoInicio").send_keys(now.strftime('01/%m/%Y'))
driver.find_element_by_id("strPeriodoFim").send_keys(now.strftime('10/%m/%Y'))
driver.find_element_by_id('checkAll').click()
driver.find_element_by_id('submit_xls').click()

# RENOMEAR OS RELATORIOS CAPTURADOS ANTES DA TRANSFORMACAO FINAL
arquivo = pathcoleta + 'tnt_' + now.strftime('%d%m%y') + '.xls'
files = glob.glob(pathcoleta + 'report.xls')

continua = True
while continua:
    if files == '[]' : break
    files = glob.glob(pathcoleta + 'report.xls')
    time.sleep(0.1)
    for file in files:
        os.rename(file, arquivo)
        continua = False

# FECHA O SITE E ENCERRA TUDO PARA ENVIO DO EMAIL
time.sleep(2)
driver.quit()

#ENVIA EMAIL DO RELATORIO FINAL PARA OS DESTINATARIOS DO PROJETO
email = 'rodrigo.matos@mw.far.br'
send_to_email = ['andrea.ferraz@mw.far.br']
cc = ['rodrigo.matos@mw.far.br','rafael.augusto@vdmoplog.com.br']
subject = 'Relatorio Diario de Fretes (Transportadora TNT)'
message = 'Ola! Segue o relatorio extraido do portal da transportadora em formato XLS, contendo todos os CNPJs. Att., Rodrigo Matos.'
file_location = arquivo
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
