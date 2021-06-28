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

pidprofilefile = r'z:\ldbot\pid\profile.pid'
pidprofile = str(os.getpid())
open(pidprofilefile, 'w').write(pidprofile)
print('Assistente PCD Perfil em execucao ...')

###-- busca na tabela profile.xlsx o nome a ser cadastrado no perfil do novo ZAP
db2 = r'z:\ldbot\db\CP' + str(dpzvalue) + 'profile.xlsx'
profile = pd.read_excel(db2, sheet_name='profileTemp', index=False)
nome = profile['NAME']
status = profile['STATUS']

time.sleep(5)
send_keys('^o') ### abre o menu SETTINGS do ZAP
time.sleep(3)
pyautogui.click(x=170, y=170) ### configura o perfil do usuario ZAP
time.sleep(3)
pyautogui.click(x=145, y=530) ### about
time.sleep(3)
pyautogui.click(x=100, y=225) ### SET About to
time.sleep(3)
pyautogui.write(status[0][:130] + '.' + str(random.randint(111111111,999999999)), interval=0.2) ### ativa uma mensagem de humor padrao
time.sleep(3)
pyautogui.click(x=1300, y=722) ### SAVE
time.sleep(3)
pyautogui.click(x=40, y=75) ### VOLTAR
time.sleep(3)
pyautogui.click(x=730, y=275) ### TROCAR FOTO
time.sleep(3)

eloc = None ###-- CLICA NO BOTAO GALERIA DE IMAGENS
while eloc == None:
	eloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\galpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
pyautogui.click(eloc)
# pyautogui.click(x=410, y=730) ### ACESSO A GALERIA DE IMAGENS

time.sleep(3)
pyautogui.click(x=120, y=225) ### ALL PHOTOS
time.sleep(3)
pyautogui.click(x=80, y=250) ### LAST WEEK
time.sleep(3)
pyautogui.click(x=1140, y=730) ### DONE
time.sleep(3)
pyautogui.click(x=40, y=75) ### VOLTAR AO CHAT 3
time.sleep(3)
pyautogui.click(x=40, y=75) ### VOLTAR AO CHAT 2
time.sleep(3)
pyautogui.click(x=40, y=75) ### VOLTAR AO CHAT 1
time.sleep(3)

pymsgbox.alert('NOVO PERFIL DE CONTA CRIADO!',timeout=2000)  ### avisa que a configuracao inicial fora concluida
subprocess.Popen(['python', r'z:\ldbot\disparazap.py'])
print('Disparazap recarregado com sucesso!')
time.sleep(2)
subprocess.Popen(['python', r'z:\ldbot\ban.py'])
print('PCD Banimento recarregado com sucesso!')

sys.exit(0)
