### ld\ldbotv3 ### updated 11.maio.2020 cadastra um novo telefone via OTP

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

pyautogui.FAILSAFE = False  # impede o mouse de provocar um erro de seguranca e abortar a operacao em qq parte do script

pidotpfile = r'z:\ldbot\pid\otp.pid'
pidotp = str(os.getpid())
open(pidotpfile, 'w').write(pidotp)
print('Assistente PCD troca de chip OTP em execucao ...')

###-- busca na tabela profile.xlsx o nome a ser cadastrado no perfil do novo ZAP
db2 = r'z:\ldbot\db\CP' + str(dpzvalue) + 'profile.xlsx'
profile = pd.read_excel(db2, sheet_name='profileTemp', index=False)
nome = profile['NAME']

db = r'z:\ldbot\db\CP' + str(dpzvalue) + 'otp.xlsx'
chip = pd.read_excel(db, sheet_name='otpTemp', index=False)
id1 = chip['ID']
idtit = chip['IDTIT']
cel = chip['FONE']
otp = chip['OTP']
uso = chip['USO']
bloq = chip['BLOQ']
serial = chip['SERIAL']

emailTo = ['ramonunix@bsd.com.br']
emailCc = ['rmatos.tec@gmail.com']

###-- carrega a tabela OTP para dentro do arquivo CSV --###
chipfile = r'z:\ldbot\db\draft.csv'
open(chipfile, 'w').write('FONE,OTP,SERIAL,ID,IDTIT,USO,BLOQ\n')
for i in cel.index:
    open(chipfile, 'a').write(str(cel[i]) + ',' + str(otp[i]) + ',' + str(serial[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(uso[i]) + ',' + str(bloq[i]) +'\n')

###-- extrai do arquivo CSV o proximo chip a ser usado que nao esteja bloqueado o em uso --###
chipcel = open(chipfile, 'r').readlines()
k = 0
m = 0
for i in chipcel:
    if i[0:1] != '0':
        k = 1
    else:
        k = 0

    if i[-4:] == '0,0\n': ### trata chip novo e okay
        print('cel: ' + i[0:10+k] + ' | otp: ' + i[11+k:17+k] + ' | serial: ' + i[18+k:23+k])
        chipold = i
        chipnew = i[0:-4] + '1,0\n'
        celnew = i[0:10+k]
        otpnew = i[11+k:17+k]
        sql =   '''
                update otp set uso = 1 where id = ''' + str(id1[m]) + '''
                ''' 
        break
    elif i[-4:] == '1,0\n': ### trata chip em uso, mas desbloqueado e okay
        chipold = ''
        chipnew = ''
        pymsgbox.alert('CHIP: ' + i[0:10+k] + ' | SERIAL: ' + i[18+k:23+k] + ' em uso no momento!',timeout=2500)
        break
    else: ### varre a lista ate encontrar um chip disponivel para troca que nao esteja em uso e nem bloqueado
        chipold = i
        chipnew = i
    m = m + 1
    
if chipold == chipnew and chipold[-4:] == '1,1\n':
    arquivo = db
    time = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
    email = 'rmatos.tec@gmail.com'
    send_to_email = emailTo #################
    cc = emailCc ############################
    subject = 'FAVOR ATIVAR NOVOS CHIPS | ' + str(time)
    message = 'Seus CHIPS acabaram!!! Favor recarregar o banco de dados OTP para continuar enviando.\r\n\r\nAtt., Equipe Disparazap 2020.'
    password = 'juizdefora2@'
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = u'Robo CHIP <rmatos.tec@gmail.com>'
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
    pyautogui.press('esc')
    pymsgbox.alert('SEUS CHIPS ACABARAM!\n\nRECARREGUE O BANCO OTP PARA CONTINUAR!',timeout=3000)
    print('Email enviado com sucesso!')
    sys.exit(0)

###-- reescreve a tabela CHIP com os dados do chip em uso logo apos a troca --###
open(chipfile, 'w').write('')
for i in chipcel:
    if i == chipold:
        open(chipfile, 'a').write(chipnew)
    else:
        open(chipfile, 'a').write(i)

###-- converte a tabela CHIP em OTP.xlsx com os dados da chipcel --###
convert = pd.read_csv (chipfile)
convert.to_excel (db, sheet_name='otpTemp', engine='xlsxwriter', index = None, header=True)

###-- grava as informacoes do chip utilizado no banco de dados db         
cursor.execute(sql)
conn.commit()

time.sleep(5)
pyautogui.click(x=700, y=610) ### agree and continue
time.sleep(2)
pyautogui.doubleClick(x=560, y=250) ### codigo pais
time.sleep(2)
pyautogui.write('55', interval=0.2) ### cadastra o codigo do pais
time.sleep(2)
pyautogui.write(celnew, interval=0.2) ### cadastra o telefone a ser usado na nova versao do zap
time.sleep(2)
pyautogui.doubleClick(x=680, y=700) ### next
time.sleep(2)

###-- aguarda a tela de confirmacao do numero informado para o otp
zloc = None
while zloc == None:
    zloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\otpinfopix1.png', grayscale=False, minSearchTime=3, confidence=.7)
time.sleep(2)

pyautogui.doubleClick(x=975, y=485) ### phone okay
time.sleep(2)
pyautogui.write(otpnew, interval=0.2) ### informa o OTP no campo destinado a digitacao
time.sleep(2)
pyautogui.doubleClick(x=775, y=455) ### phone NEXT
time.sleep(2)

###-- aguarda a tela do OTP passar adiante, caso falhe na ativacao
gloc = None
gloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sixdigitpix1.png', grayscale=False, minSearchTime=5, confidence=.7)
while gloc != None:
    gloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sixdigitpix1.png', grayscale=False, minSearchTime=3, confidence=.7)
    time.sleep(2)
time.sleep(2)

###-- verifica se a tela de cadastro de perfil passou direto sem pedir bkp
eloc = None
floc = None
eloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\profileinfopix1.png', grayscale=False, minSearchTime=5, confidence=.7)
floc = pyautogui.locateOnScreen(r'z:\ldbot\pix\profileinfopix2.png', grayscale=False, minSearchTime=5, confidence=.7)
if eloc != None or floc != None:
    time.sleep(2)
    pyautogui.write(nome[0][:18] + '.' + str(random.randint(111111,999999)), interval=0.2) ### preenche o nome do responsavel pelo numero do zap (matriz)
    time.sleep(2)
    pyautogui.doubleClick(x=680, y=720) ### NOME PROFILE NEXT
    time.sleep(2)
    ###-- aguarda a tela de inicializacao do novo chip passar adiante
    nloc = None
    nloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\initotppix1.png', grayscale=False, minSearchTime=5, confidence=.7)
    while nloc != None:
        nloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\initotppix1.png', grayscale=False, minSearchTime=3, confidence=.7)
    time.sleep(3)

    pymsgbox.alert('PRONTO PARA INICIAR OS AJUSTES DO PERFIL DO USUARIO!',timeout=2000)  ### avisa da configuracao inicial concluida... indo para o perfil

    subprocess.Popen(['python', r'z:\ldbot\profile.py'])

    sys.exit(0)

time.sleep(3)

zloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\skippix1.png', grayscale=False, minSearchTime=3, confidence=.7)
if zloc != None:
    time.sleep(2)
    pyautogui.doubleClick(x=770, y=455) ###  SKIP Permission do google accounts backup e clica em SKIP

aloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\bkpfoundpix1.png', grayscale=False, minSearchTime=5, confidence=.7)
if aloc != None:
    time.sleep(2)
    pyautogui.doubleClick(x=790, y=715) ### phone BACKUP FOUND SKIP

bloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\bkpfoundpix2.png', grayscale=False, minSearchTime=5, confidence=.7)
if bloc != None:
    time.sleep(2)
    pyautogui.doubleClick(x=935, y=445) ### phone BACKUP FOUND SKIP CONFIRMATION

###-- verifica se a tela de cadastro de perfil apareceu
tloc = None
uloc = None
tloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\profileinfopix1.png', grayscale=False, minSearchTime=5, confidence=.7)
uloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\profileinfopix2.png', grayscale=False, minSearchTime=5, confidence=.7)
if tloc != None or uloc != None:
    time.sleep(2)
    pyautogui.write(nome[0][:18] + '.' + str(random.randint(111111,999999)), interval=0.2) ### preenche o nome do responsavel pelo numero do zap (matriz)
    time.sleep(2)
    pyautogui.doubleClick(x=680, y=720) ### NOME PROFILE NEXT
    time.sleep(2)

###-- aguarda a tela de inicializacao do novo chip passar adiante
nloc = None
nloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\initotppix1.png', grayscale=False, minSearchTime=5, confidence=.7)
while nloc != None:
    nloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\initotppix1.png', grayscale=False, minSearchTime=3, confidence=.7)
time.sleep(3)

pymsgbox.alert('PRONTO PARA INICIAR OS AJUSTES DO PERFIL DO USUARIO!',timeout=2000)  ### avisa da configuracao inicial concluida... indo para o perfil

subprocess.Popen(['python', r'z:\ldbot\profile.py'])

sys.exit(0)