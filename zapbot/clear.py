### ld\ldbotv3 ### updated 11.maio.2020 apagar dados do zap apos o banimento

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

pidclearfile = r'z:\ldbot\pid\clear.pid'
pidclear = str(os.getpid())
open(pidclearfile, 'w').write(pidclear)
print('Assistente PCD limpeza chip banido em execucao ...')

######subprocess.Popen(['rasdial','3g','/disconnect']) ### desabilita a interface de rede
######pymsgbox.alert('MODEM 3G DESCONECTADO',timeout=2000)  ### avisa que o modem fora desconectado

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

###-- extrai do arquivo CSV o chip em uso que precisa ser bloqueado --###
chipcel = open(chipfile, 'r').readlines()
k = 0
m = 0
for i in chipcel:
    if i[0:1] != '0':
        k = 1
    else:
        k = 0

    if i[-4:] == '1,0\n': ### trata chip em uso para bloquea-lo
        print('cel: ' + i[0:10+k] + ' | otp: ' + i[11+k:17+k] + ' | serial: ' + i[18+k:23+k])
        chipold = i
        chipnew = i[0:-4] + '1,1\n'
        sql =   '''
                update otp set uso = 1, bloq = 1 where id = ''' + str(id1[m]) + '''
                ''' 
        break
    else: ### varre a lista ate encontrar o chip que esta em uso para bloquea-lo
        chipold = i
        chipnew = i
    m = m + 1

if chipold == chipnew:
    pymsgbox.alert('NENHUM CHIP EM USO DETECTADO... TODOS NOVOS!',timeout=2000)
        
###-- reescreve a tabela CHIP com os dados do chip logo apos seu bloqueio --###
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

###-- abertura de SETTINGS para zerar o ZAP --###
subprocess.Popen([r'C:\ChangZhi\LDPlayer\dnconsole.exe', 'launchex', '--index', '0', '--packagename', 'com.android.settings']) ### acessa os ajustes
time.sleep(3)

vloc = None
while vloc == None: ### clica no APPS
	vloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\appspix1.png', grayscale=False, minSearchTime=3, confidence=.8)
pyautogui.click(vloc)
time.sleep(2)
zloc = None
while zloc == None: ### clica no ZAP
	zloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\zappix1.png', grayscale=False, minSearchTime=3, confidence=.8)
pyautogui.click(zloc)
time.sleep(2)

pyautogui.click(x=1045, y=550) ### clear data
time.sleep(3)
pyautogui.click(x=975, y=475) ### clear data okay
time.sleep(3)
pyautogui.click(x=300, y=230) ### force stop
time.sleep(3)
pyautogui.click(x=975, y=470) ### force stop okay
time.sleep(2)
pymsgbox.alert('PROCESSO DE LIMPEZA DO ZAP CONCLUIDA... PRONTO PARA INSTALAR NOVO CHIP!',timeout=2000)  ### avisa do zeramento do ZAP
time.sleep(2)

subprocess.Popen(['ipconfig', '/flushdns'])  ### limpa o cache DNS
subprocess.Popen([r'C:\Program Files\CCleaner\CCleaner64.exe', '/auto'])  ### executa o ccleaner para limpeza geral
time.sleep(2)
######subprocess.Popen(['rasdial','3g']) ### habilita modem 3g
######pymsgbox.alert('MODEM 3G CONECTADO A REDE!!!',timeout=2000)  ### avisa que o modem fora conectado

subprocess.Popen([r'C:\ChangZhi\LDPlayer\dnconsole.exe', 'runapp', '--index','0','--packagename','com.whatsapp'])
time.sleep(3)
subprocess.Popen(['python', r'z:\ldbot\otp.py'])

sys.exit(0)