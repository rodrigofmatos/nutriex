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
dpzvalue = open(dpzfile, 'r').readline()[:-1]
dpzlote = r'z:\ldbot\db\cplotes.dpz'
dpzvalote = open(dpzlote, 'r').readlines()
mdb = r'z:\ldbot\db\db.accdb'
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + mdb + ';')
cursor = conn.cursor()
maquina = int(dpzvalue)

# --- carrega o banco de dados para a memoria ---#
db = r'z:\ldbot\db\CP' + str(dpzvalue) + 'tit.xlsx'
db3 = r'z:\ldbot\db\CP' + str(dpzvalue) + 'msg.xlsx'

tit = pd.read_excel(db, sheet_name='titTemp', index=False)
titulo = str(tit['ID']) + '_' + str(tit['VM']) + '_' + str(tit['CLIENTE']) + '_' + str(tit['CAMPANHA'])
msg = pd.read_excel(db3, sheet_name='msgTemp', index=False)
anuncio = msg['MENSAGEM']

emailTo = ['ramonunix@bsd.com.br']
emailCc = ['rmatos.tec@gmail.com']
z = 0

for lote in dpzvalote:
    db2 = r'z:\ldbot\db\CP' + str(dpzvalue) + 'LT' + lote[:-1] + '.xlsx'
    if os.path.exists(db2):
        pass
    else:
        print(str(db2) + ' nao encontrado!!!')
        pymsgbox.alert('Arquivo: ' + str(db2) + ' não foi encontrado!',timeout=2000)
        sys.exit(0)
    cel = pd.read_excel(db2, sheet_name='CP' + str(dpzvalue) + 'LT' + lote[:-1], index=False)
    celular = cel['FONE_MESCLADO']
    id1 = cel['ID']
    idtit = cel['IDTIT']
    controlid = cel['CONTROLID']
    dthr = cel['DTHR']
    env = cel['ENVIADO']
    if len(celular) == 0: 
        pymsgbox.alert('Não existem contatos para o disparo. Planilha VAZIA!',timeout=2000)
        pidbanfile = r'z:\ldbot\pid\ban.pid'
        pidban = int(open(pidbanfile, 'r').read())    
        try:
            os.kill(pidban, signal.SIGTERM)
        except OSError:
            pymsgbox.alert('PCD BANIMENTO ENCERRADO!',timeout=2000)

        sys.exit(0)
    ###-- carrega a tabela CEL para dentro do arquivo LIST.CSV --###
    chipfile = r'z:\ldbot\db\list.csv'
    open(chipfile, 'w').write('FONE_MESCLADO,ID,IDTIT,CONTROLID,DTHR,ENVIADO\n')
    varold = 0 ###-- armazena a quantidade de celulares enviados
    varnew = 0 ###-- armazena a quantidade de celulares novos e ainda nao enviados

    for i in celular.index:
        if str(env[i]).upper() == 'Y':
            enviado = 'Y'
            data = str(dthr[i])
            varold = varold + 1
        else:
            enviado = 'N'
            data = str(datetime.datetime.now())
            varnew = varnew + 1
        open(chipfile, 'a').write(str(celular[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(controlid[i]) + ',' + data + ',' + enviado + '\n')

    if varold == len(celular):
        pidbanfile = r'z:\ldbot\pid\ban.pid'
        pidban = int(open(pidbanfile, 'r').read())    
        try:
            os.kill(pidban, signal.SIGTERM)
        except OSError:
            pymsgbox.alert('PCD BANIMENTO ENCERRADO!',timeout=2000)

        print('Assistente PCD Banimento encerrado!')
        pymsgbox.alert('VOCE ENVIOU TODOS OS SEUS ' + str(len(dpzvalote)*varold) + ' CONTATOS!',timeout=2000)
        break
    else:
        pymsgbox.alert('O LOTE ' + lote[:-1] + ' AINDA TEM ' + str(varnew) +  ' CONTATOS PARA ENVIO.',timeout=2000)
            
    # --- elimina caracteres especiais do nome do arquivo antes de cria-lo  ---#
    tit = re.sub("[ç]", "c", str(titulo[int(maquina) - 1]))
    tit = re.sub("[ãáâ]", "a", str(titulo[int(maquina) - 1]))
    tit = re.sub("[êé]", "e", str(titulo[int(maquina) - 1]))
    tit = re.sub("[óôõ]", "o", str(titulo[int(maquina) - 1]))
    tit = re.sub("[í]", "i", str(titulo[int(maquina) - 1]))
    tit = re.sub("[ú]", "u", str(titulo[int(maquina) - 1]))
    tit = re.sub(r'[^a-zA-Z0-9_]', r'', str(titulo[int(maquina) - 1]))

    # --- ativa o arquivo de log para gravacao dos eventos e extracao do relatorio ao final da transmissao ---#
    now = datetime.datetime.now()
    now_str = now.strftime("%Y%m%d-%Hh%Mm")

    tempoespera = 1  ### tempo de espera da variavel sleep definido em segundos
    bloco = 0  # variavel que define a varredura de todos os celulares que serao utilizados ate igualar a variavel repeticao
    a = 0  # variavel que determina o numero do anuncio que esta em vigor
    repeticao = len(celular) ### usa o total de celulares cadastrados no Excel para fazer o loop entre dos envios
    intervalo = 5 ### intervalo padrao do bloco (de 5 em 5)
    limitek = intervalo ### limite determinado pelo intervalo ou fracao dele, caso a quantidade de celulares nao seja multipla de 5
    saldo = (int(repeticao / intervalo) * intervalo) ### usado para determinar o multiplo inteiro mais proximo do total de celulares

    # ==================== BLOCO 001 ====================#
    # --- inicio da repeticao em massa --- #
    i = 0

    open(chipfile, 'w').write('FONE_MESCLADO,ID,IDTIT,CONTROLID,DTHR,ENVIADO\n')
    while bloco < repeticao: ### inicia o envio em massa
     
        k = 0

        if bloco == saldo: ### quando ele detecta que o contato ativo se igualou ao maximo multiplo, limitek passa a ter o valor como o resto da operacao (1, 2, 3 ou 4)
            limitek = repeticao % intervalo
            
        while k < limitek: #--- bloco de repeticao de 5 em 5 ou ate que o limitek seja atingido (nao necessariamente multiplo de 5)    
            sql =   '''
                    update dbtest set enviado = 'Y' where id = ''' + str(id1[i]) + '''
                '''            
            if env[i] == 'N':
                enviado = 'N'
                data = str(datetime.datetime.now())
                chipold = str(celular[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(controlid[i]) + ',' + str(dthr[i]) + ',' + str(env[i]) + '\n'
                chipnew = str(celular[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(controlid[i]) + ',' + data + ',Y\n'
                celnew = str(celular[i])
                print('celular: ' + str(celular[i]) + ' | id: ' + str(id1[i]) + ' | idtit: ' + str(idtit[i]) + ' | controle: ' + str(controlid[i]) + ' | data: ' + data + ' | enviado?: Y')
                open(chipfile, 'a').write(chipnew)
                i = i + 1
                k = k + 1
                bloco = bloco + 1
                cursor.execute(sql)
                conn.commit()
            else:
                chipold = str(celular[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(controlid[i]) + ',' + str(dthr[i]) + ',' + str(env[i]) + '\n'
                chipnew = str(celular[i]) + ',' + str(id1[i]) + ',' + str(idtit[i]) + ',' + str(controlid[i]) + ',' + str(dthr[i]) + ',' + str(env[i]) + '\n'
                print('celular: ' + str(celular[i]) + ' | id: ' + str(id1[i]) + ' | idtit: ' + str(idtit[i]) + ' | controle: ' + str(controlid[i]) + ' | data: ' + data + ' | enviado?: ' + str(env[i]))
                open(chipfile, 'a').write(chipold)
                i = i + 1
                k = k + 1
                bloco = bloco + 1                
                cursor.execute(sql)
                conn.commit()

            if chipold == chipnew:
                pymsgbox.alert('SUA LISTA DE ENVIO CHEGOU AO FIM, PARABENS!',timeout=1500)
                break 
        ###-- converte a tabela CEL em CEL.xlsx com os dados da chipcel --###
        z = z + k
        print('i: ' + str(i))
        print('k: ' + str(k))
        print('z: ' + str(z))
        print('bloco:' + str(bloco))

        convert = pd.read_csv (chipfile)
        convert.to_excel (db2, sheet_name='CP' + str(dpzvalue) + 'LT' + lote[:-1], engine='xlsxwriter', index = None, header=True)
     
        # if a < len(anuncio)-1: ### define o numero do anuncio que sera enviado (um diferente por bloco)
        #     a = a + 1
        # else:
        #     a = 0
        
        # if (z % (20)) == 0: ###--- esvaziamento de cache a cada 40 contatos
        #     # subprocess.Popen(['ipconfig', '/flushdns'])  ### limpa o cache DNS
        #     # subprocess.Popen([r'C:\Program Files\CCleaner\CCleaner64.exe', '/auto'])  ### executa o ccleaner para limpeza geral
        #     pymsgbox.alert('EXECUTANDO LIMPEZA GERAL. FLUSH DNS E CCLEANER...',timeout=500)  ### avisa sobre a limpeza geral periodica
            
        # if (z % (30)) == 0: ###--- limpeza de cache e DNS a cada 60 contatos
        #     pymsgbox.alert('CACHE DO ZAP ESVAZIADO!!!',timeout=500)  ### avisa que o cache do zap fora esvaziado com sucesso
  
    # ==================== BLOCO 010 ====================#
    if (bloco % intervalo) == 1: ### valida o caso de o envio ser para apenas um contato no caso o total de contatos nao ser multiplo de 5 com resto 1
        pymsgbox.alert('O PROGRAMA SERA ENCERRADO PORQUE CHEGOU AO FIM!',timeout=1000)
        #################### ENCERRA O SCRIPT E AS APLICACOES OS RESPECTIVOS PCDs ############################
        pidbanfile = r'z:\ldbot\pid\ban.pid'
        pidban = int(open(pidbanfile, 'r').read())    
        try:
            os.kill(pidban, signal.SIGTERM)
        except OSError:
            pymsgbox.alert('PCD BANIMENTO JA ESTAVA ENCERRADO!',timeout=2000)

        print('Assistente PCD Banimento encerrado!')
        pymsgbox.alert('Programa encerrado! E-mail enviado, log gerado e fotos tiradas...',timeout=4000)  ### encerra o envio das mensagens da campanha ativa
        break

    #--- envio do email finalizando as entregas mais o arquivo de log de toda a operacao ---#
    arquivo = r'z:\ldbot\log\{}-{}.log'.format(now_str, tit)
    tempo = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')

#################### ENCERRA O SCRIPT E AS APLICACOES OS RESPECTIVOS PCDs ############################
pidbanfile = r'z:\ldbot\pid\ban.pid'
pidban = int(open(pidbanfile, 'r').read())    
try:
    os.kill(pidban, signal.SIGTERM)
except OSError:
    pymsgbox.alert('PCD BANIMENTO ENCERRADO!',timeout=1000)

print('Assistente PCD Banimento encerrado!')
pymsgbox.alert('TESTE REALIZADO COM SUCESSO!',timeout=4000)  ### encerra o envio das mensagens da campanha ativa
sys.exit(0)