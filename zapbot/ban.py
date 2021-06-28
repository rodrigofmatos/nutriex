### ld\ldbotv1 ### updated 21.abr.2020 - leitura do arquivo pid e deteccao do banimento com envio de email
### ld\ldbotv3 ### updated 08.maio.2020 - correcao dos tempos de deteccao da tela de banimento
import pymsgbox
import random
from pywinauto.keyboard import send_keys
from pywinauto.mouse import scroll
from pytz import timezone
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
import logging
import datetime
import numpy as np
import imutils
import cv2
import subprocess
import pyautogui
import pyperclip
import time
import pandas as pd
import os
import signal
import sys
from io import StringIO
import pyodbc

dpzfile = r'z:\ldbot\db\cpativa.dpz'
dpzvalue = str(open(dpzfile, 'r').readline())[:-1]
mdb = r'z:\ldbot\db\db.accdb'
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + mdb + ';')
cursor = conn.cursor()

pidzapfile = r'z:\ldbot\pid\disparazap.pid'
pidbanfile = r'z:\ldbot\pid\ban.pid'
pidban = str(os.getpid())
open(pidbanfile, 'w').write(pidban)
emailTo = ['ramonunix@bsd.com.br']
emailCc = ['rmatos.tec@gmail.com']
        
print('Assistente PCD Banimento em execucao ...')
#--- ativa o arquivo de log para gravacao dos eventos e extracao do relatorio ao final da transmissao ---#
now = datetime.datetime.now()
now_str = now.strftime("%Y%m%d-%Hh%Mm")

#==================== BLOCO 999 ====================#
#--- localizador de imagens ---# VERIFICACAO DE BANIMENTO
bloc = None
i = 0

while bloc == None:
    bloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\banpix1.png', grayscale=False, minSearchTime=5, confidence=.7)
    time.sleep(5)
    if bloc != None:
        pyautogui.click(x=975, y=475) ### botao OK para fechar a tela do banimento
        bloc = None
        pidzap = int(open(pidzapfile, 'r').read())
        os.kill(pidzap, signal.SIGTERM)
        arquivo = r'z:\ldbot\log\{}-{}.log'.format(now_str,'BANIMENTO')
        time = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
        email = 'rmatos.tec@gmail.com'
        send_to_email = emailTo #################
        cc = emailCc ############################
        subject = 'BANIMENTO | ' + str(time)
        message = 'Segue o relatorio emitido pelo robo DISPARAZAP\r\ncom detalhes do BANIMENTO.\r\n\r\nAtt., Equipe Disparazap 2020.'
        password = 'juizdefora2@'
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = u'Robo BANIMENTO <rmatos.tec@gmail.com>'
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
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, send_to_email+cc, text)
        server.quit()
        
        pymsgbox.alert('<<<< BANIMENTO DO CHIP >>>>\r\n\r\nCONFIRA O RELATÓRIO EM:\r\n\r\n(PASTA: log >>> {}-{}.log'.format(now_str,'BANIMENTO') + ')',timeout=3000)  ### avisa houve banimento
        subprocess.Popen(['python', r'z:\ldbot\clear.py'])
         
        sys.exit(0)
