import pandas as pd
import time
import sys
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
import ntpath
from pywinauto.keyboard import send_keys
import subprocess
import pyautogui
from selenium.webdriver.common.keys import Keys
from selenium import webdriver

now = datetime.datetime.now().strftime("%H")

if int(now) >= 7 and int(now) <= 20:
    site = 'http://192.168.1.15:81/vaptvupt/editais.aspx'
    driver = webdriver.Chrome(r'f:\bot\rpa017\chromedriver.exe')
    driver.get(site)
    time.sleep(3)

    titulo = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    start_time = time.time() ###-- inicia a temporizacao do codigo
    emailTo = ['rodrigo.matos@nutriex.com.br']
    emailCc = ['dl21@dldistribuidora.net.br']
    db_result= r'f:\bot\rpa017\pregao_{}.html'.format(titulo)
    time.sleep(3)

    # subprocess.Popen([r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', site]) 
    send_keys('^s')
    time.sleep(1)
    pyautogui.typewrite(db_result)
    time.sleep(1)
    send_keys('{ENTER}')
    time.sleep(8)
    driver.quit()
    
    arquivo = db_result
    tempo = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
    email = 'robot@nutriex.com.br'
    send_to_email = emailTo #################
    cc = emailCc ############################
    subject = "RELATORIO MONITORAMENTO DIARIO PREGAO | " + str(tempo)
    message = "Segue o relatório diário contendo o pregão obtido hoje.\r\n\r\n<Tempo total de processamento: %.1f segundos>" % (time.time() - start_time) + "\r\n\r\nAtt., Equipe Tecnologia 2020."
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

sys.exit(0)
