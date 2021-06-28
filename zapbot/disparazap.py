### ld\ldbotv2 ### updated 26.abr.2020  limpeza do codigo do projeto ldplayer
### ld\ldbotv2 ### updated 27.abr.2020  imagens redimensionadas e comandos do ld trocados
### ld\ldbotv2 ### updated 28.abr.2020  testes e ajustes no codigo
### ld\ldbotv2 ### updated 29.abr.2020  problemas nos ajustes de foco entre prompt de comando Win7 e LDPlayer\dnconsole
### ld\ldbotv2 ### updated 30.abr.2020  prova final para liberar em producao
### ld\ldbotv2 ### updated 02.maio.2020 comando que transfere para area de transferencia pyperclip apresentando instabilidade (tkinter)
### ld\ldbotv2 ### updated 04.maio.2020 reuniao com Ramon 21h para definir os proximos passos - testes okay ate o disparo dos 5 primeiros
### ld\ldbotv2 ### updated 05.maio.2020 loop completo funcionando, fotos das etapas concluidas, logs ativados, envio de email e banimento
### ld\ldbotv2 ### updated 06.maio.2020 deteccao de banimento
### ld\ldbotv2 ### updated 06.maio.2020 log dos celulares enviados
### ld\ldbotv2 ### updated 06.maio.2020 promover limpeza da tela inicial
### ld\ldbotv2 ### updated 06.maio.2020 remover a mensagem de finalizacao (pop-up)
### ld\ldbotv2 ### updated 07.maio.2020 limpeza de cache do zap
### ld\ldbotv3 ### updated 08.maio.2020 testar velocidade de envio aleatório entre os blocos (entre 1 e 9 segundos de espera)
### ld\ldbotv3 ### updated 11.maio.2020 configurar cadastro de chip com otp pós banimento
### ld\ldbotv3 ### updated 11.maio.2020 personalizar status do chip novo
### ld\ldbotv3 ### updated 15.maio.2020 regulagem dos tempos
### ld\ldbotv3 ### updated 16.maio.2020 validacao do processo otp com base na tabela Excel
### ld\ldbotv3 ### updated 18.maio.2020 validacao do processo celulares com base na tabela Excel
### ld\ldbotv3 ### updated 19.maio.2020 validacao do processo real na maq do Ramon
### ld\ldbotv3 ### updated 23.maio.2020 troca do processo de validacao independentemente da maq utilizada
### ld\ldbotv3 ### updated 24.maio.2020 abertura de settings antes do zap + alternar por F2 + iniciar busca pela lupa
### ld\ldbotv3 ### updated 25.maio.2020 ajustes do perfil do novo ZAP a partir da planilha PROFILE.xlsx
### ld\ldbotv3 ### updated 28.maio.2020 ajustes finos com adicao de mais controles de erros nas rotinas do sistema
### ld\ldbotv3 ### updated 29.maio.2020 automatizacao da abertura do LD + checagem de fullscreen + correcao zap profile
### ld\ldbotv3 ### updated 03.junho.2020 correcao do erro detectado durante os testes: aumentar o tempo do SHARE depois do MKT
### ld\ldbotv3 ### updated 04.junho.2020 ao inserir novo OTP, criar ajuste inicial de profile antes de tudo, caso nao surja msg de bkp
### ld\ldbotv3 ### updated 05.junho.2020 corrigido a deteccao do segundo banimento após a troca do chip (ban.py nao era chamado novamente)
### ld\ldbotv3 ### updated 05.junho.2020 dica 2: alternar nome acrescentando codigo aleatorio de 7 digitos
### ld\ldbotv3 ### updated 05.junho.2020 dica 2: alternar about acrescentando codigo aleatorio de 9 digitos
### ld\ldbotv3 ### updated 05.junho.2020 erro 2: ele nao trocou a foto do perfil durante o profile... reconhecer a imagem do botao
### ld\ldbotv3 ### updated 06.junho.2020 dica 3: acrescentar aos textos numeros aleatorios
### ld\ldbotv3 ### updated 06.junho.2020 dica 11: separar aba MSG do arquivo CAMPANHA.xlsx
### ld\ldbotv3 ### updated 06.junho.2020 dica 12: separar aba TIT do arquivo CAMPANHA.xlsx
### ld\ldbotv3 ### updated 06.junho.2020 dica 5: renumerar arquivos de fotos do rep com data hora
### ld\ldbotv4 ### updated 26.junho.2020 modificacao estrutural do banco para access 2010+ 64bits accdb, python 64 bits
### ld\ldbotv4 ### updated 29.junho.2020 loop entre campanhas e lotes gerados pelo Access

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
#maincontact = 'dispara zap' #'6285194924'
maincontact = '6291085505'
z = 0

# ==================== BLOCO 000 ====================#
# --- tela inicial de configuracao da resolucao do monitor para correto funcionamento da aplicacao ---#
pyautogui.FAILSAFE = False  # impede o mouse de provocar um erro de seguranca e abortar a operacao em qq parte do script
pidfile = r'z:\ldbot\pid\disparazap.pid'
pid = str(os.getpid())
open(pidfile, 'w').write(pid) ### grava o numero do processo PID do disparazap para encerra-lo ao final de tudo

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
        pyautogui.press('f1') ### vai para a tela HOME do emulador
        time.sleep(2)
        pyautogui.press('f12') ### restaurar a janela do emulador
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
 
    tempoespera = 3  ### tempo de espera da variavel sleep definido em segundos
    bloco = 0  # variavel que define a varredura de todos os celulares que serao utilizados ate igualar a variavel repeticao
    a = 0  # variavel que determina o numero do anuncio que esta em vigor
    repeticao = len(celular) ### usa o total de celulares cadastrados no Excel para fazer o loop entre dos envios
    intervalo = 5 ### intervalo padrao do bloco (de 5 em 5)
    limitek = intervalo ### limite determinado pelo intervalo ou fracao dele, caso a quantidade de celulares nao seja multipla de 5
    saldo = (int(repeticao / intervalo) * intervalo) ### usado para determinar o multiplo inteiro mais proximo do total de celulares

    # ==================== MATRIZ 003 ====================#
    ### inicia o envio com base nas condicoes da lista: Start a chat > vai comecar direto, sem limpeza
    bloc = None
    cloc = None
    bloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\startpix2.png', grayscale=False, minSearchTime=3, confidence=.8)
    cloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\startpix1.png', grayscale=False, minSearchTime=3, confidence=.8)
    time.sleep(2)
    if bloc != None or cloc != None:
        pyautogui.click(x=1275,y=75) ### opcao LUPA
        time.sleep(0.8)
        pyautogui.write(maincontact, interval=0.3) ### informa o numero do contato principal que sera a matriz a ser multiplicada
        time.sleep(3)
        pyautogui.click(x=150, y=210)
    else: ### a partir daqui, significa que tem um ou mais contatos na lista para serem apagados
        pyautogui.mouseDown(x=170,y=230) ### opcao ESCOLHE CONTATO LISTA
        time.sleep(3)
        pyautogui.mouseUp(x=170,y=230) ### opcao MARCA CONTATO NA LISTA
        pyautogui.click(x=1335, y=75) ###--- clica no MENU SELECT ALL
        time.sleep(0.8)
        rloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\selectallpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
        if rloc != None: ### aqui com certeza tem mais de um contato na lista para exclusao
            pyautogui.click(rloc) ###--- SELECT ALL
            time.sleep(0.8)
            pyautogui.click(x=1138, y=75) ###--- clica na LIXEIRA
            time.sleep(0.8)
            pyautogui.click(x=970, y=470) ###--- clica no DELETE
            time.sleep(0.8)
            pyautogui.click(x=1275,y=75) ### opcao LUPA
            time.sleep(0.8)
            pyautogui.write(maincontact, interval=0.3) ### informa o numero do contato principal que sera a matriz a ser multiplicada
            time.sleep(5)
            pyautogui.click(x=150, y=210)    
        else:
            pyautogui.click(x=900, y=75) ###--- clica FORA DA LISTA POIS TEM UM ITEM APENAS
            time.sleep(0.8)
            pyautogui.click(x=1138, y=75) ###--- clica na LIXEIRA
            time.sleep(0.8)
            pyautogui.click(x=970, y=470) ###--- clica no DELETE
            time.sleep(0.8)
            pyautogui.click(x=1275,y=75) ### opcao LUPA
            time.sleep(0.8)
            pyautogui.write(maincontact, interval=0.3) ### informa o numero do contato principal que sera a matriz a ser multiplicada
            time.sleep(5)
            pyautogui.click(x=150, y=210)         
  
    # ==================== MATRIZ 004 ====================#
    # --- localizador de imagens ---# \ldbotao TELA MATRIZ ANUNCIOS
    # --- localizador de imagens ---# vai clicar no botao clips da galeria
    zloc = None
    while zloc == None:
        zloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\clippix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.click(zloc)

    # --- localizador de imagens ---# vai clicar no botao galeria
    eloc = None
    while eloc == None:
        eloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\galpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.click(eloc)

    # --- localizador de imagens ---# vai clicar no botao Send To da biblioteca de imagens
    floc = None
    while floc == None:
        floc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendtopix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.click(x=100, y=200)

    # --- localizador de imagens ---# vai clicar no botao All Media, onde fica a imagem a ser selecionada
    gloc = None
    while gloc == None:
        gloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\allmediapix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.click(x=70, y=240)

    # --- localizador de imagens ---# vai clicar no botao do triangulo ENVIAR para mandar a mensagem e gravar na matriz
    hloc = None
    while hloc == None:
        hloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.click(hloc)	


    # ==================== MATRIZ 005 ====================#
    # --- localizador de imagens ---# \ldbotao ENCAMINHAMENTO DE MENSAGENS
    # --- localizador de imagens ---# vai clicar no botao COMPARTILHAR da imagem para ativar o compartilhamento de mensagens
    iloc = None
    time.sleep(1)
    while iloc == None:
        iloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\forwardpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
    pyautogui.mouseDown(x=iloc.left-30, y=iloc.top, button='left') ### clica e segura o mouse ate que o botao SHARE apareca
    time.sleep(7)
    pyautogui.mouseUp(x=iloc.left-30, y=iloc.top, button='left')
    time.sleep(2)
    pyautogui.click(x=1270, y=75) ### --- clica no botao SHARE na posicao absoluta, sem testar


    # time.sleep(6)
    # img = cv2.imread(u'z:\\ldbot\\pix\\banpix1.png')
    # cv2.imshow('BANIDO!!!',img)
    # cv2.moveWindow('BANIDO!!!',int(1366/4),int(768/4))
    # cv2.waitKey(4444)


    # ==================== BLOCO 001 ====================#
    # --- inicio da repeticao em massa --- #
    i = 0

    open(chipfile, 'w').write('FONE_MESCLADO,ID,IDTIT,CONTROLID,DTHR,ENVIADO\n')
    while bloco < repeticao: ### inicia o envio em massa
     
        k = 0
        # --- localizador de imagens ---# \ldbotao ENCONTRA O BOTAO DE NOVOS CONTATOS PARA INICIAR AS CONVERSAS
        bloc = None
        while bloc == None:
            bloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\chatspix1.png', grayscale=False, minSearchTime=6, confidence=.7)
        pyautogui.click(bloc)
        
        time.sleep(3)

        # --- localizador de imagens ---# \ldbotao identificar Start New Group Family
        xloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\startgrouppix1.png', grayscale=False, minSearchTime=3, confidence=.8)
        if xloc != None:
            pyautogui.click(x=1332, y=240) ### apertar o x para fechar a caixa de dialogo de criacao de um novo grupo familiar
        time.sleep(0.8)

        pyautogui.click(x=1275,y=75) ### opcao LUPA
        time.sleep(0.8)
        pyautogui.write(maincontact, interval=0.3) ### informa o numero do contato principal que sera a matriz a ser multiplicada
        time.sleep(5)
        pyautogui.click(x=150, y=210)   

        # --- localizador de imagens ---# \ldbotao ENCAMINHAMENTO DE MENSAGENS
        iloc = None
        while iloc == None:
            iloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\forwardpix1.png', grayscale=False, minSearchTime=6, confidence=.7)
        pyautogui.mouseDown(x=iloc.left-30, y=iloc.top, button='left')
        time.sleep(7)
        pyautogui.mouseUp(x=iloc.left-30, y=iloc.top, button='left')
        time.sleep(2)
        pyautogui.click(x=1270, y=75) ### --- clica no botao SHARE na posicao absoluta, sem testar

        if bloco == saldo: ### quando ele detecta que o contato ativo se igualou ao maximo multiplo, limitek passa a ter o valor como o resto da operacao (1, 2, 3 ou 4)
            limitek = repeticao % intervalo
        
        # --- localizador de imagens ---# vai clicar no botao lupa para localizar o primeiro registro da repeticao dos 5 celulares
        kloc = None
        while kloc == None:
            kloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\lupapix1.png', grayscale=False, minSearchTime=6, confidence=.7)
            
        while k < limitek: #--- bloco de repeticao de 5 em 5 ou ate que o limitek seja atingido (nao necessariamente multiplo de 5)    
            sql =   '''
                    update cel set enviado = 'Y' where id = ''' + str(id1[i]) + '''
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

            ###-- continua o preenchimento dos telefones da lista
            pyautogui.click(kloc)
            time.sleep(0.8)
            pyautogui.write(str(celnew), interval=0.3) ### insere o valor do contato obtido da matriz de celulares do Excel
            time.sleep(1.1)
            pyautogui.click(x=200, y=170)

            # --- localizador de imagens ---# vai clicar no botao X para fechar a barra de pesquisa e permitir a digitacao de um novo contato
            mloc = None
            while mloc == None:
                mloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\closepix1.png', grayscale=False, minSearchTime=6, confidence=.7)
            pyautogui.click(mloc)

        ###-- converte a tabela CEL em CEL.xlsx com os dados da chipcel --###
        z = z + k
        print('i: ' + str(i))
        print('k: ' + str(k))
        print('z: ' + str(z))
        print('bloco:' + str(bloco))

        convert = pd.read_csv (chipfile)
        convert.to_excel (db2, sheet_name='CP' + str(dpzvalue) + 'LT' + lote[:-1], engine='xlsxwriter', index = None, header=True)

        # --- localizador de imagens ---# vai clicar no botao para ENVIAR o bloco de contatos compartilhados
        nloc = None
        while nloc == None:
            nloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendpix2.png', grayscale=False, minSearchTime=6, confidence=.7)
        pyautogui.click(nloc)

        time.sleep(0.8)
        # --- localizador de imagens ---# apenas avalia se a pagina anterior foi fechada e se o programa encontra-se na pagina de envio do anuncio
        oloc = None
        while oloc == None:
            oloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendpix1.png', grayscale=False, minSearchTime=6, confidence=.7)

        ##### ---- anuncio que sera colado na mensagem ---####
        pyperclip.copy((str(anuncio[a] + '.' + str(random.randint(111111111,999999999)))))
        time.sleep(4)
        pyautogui.press('win', presses=2, interval=0.6) ### alterna rapidamente entre o Windows e o Emulador, para permitir que o processo de Copiar e Colar funcione
        time.sleep(0.8)
        pyautogui.click(x=140, y=700)
        time.sleep(3)
        send_keys('^v')
        time.sleep(random.randint(5,8)) ### gera um tempo aleatorio para envio das mensagens

        if a < len(anuncio)-1: ### define o numero do anuncio que sera enviado (um diferente por bloco)
            a = a + 1
        else:
            a = 0

        # --- localizador de imagens ---# define qual o formato do botao que vai aparecer antes de enviar o anuncio para os contatos escolhidos
        qloc = None
        ploc = None
        while True:
            qloc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendpix3.png', grayscale=False, minSearchTime=6, confidence=.6)
            ploc = pyautogui.locateOnScreen(r'z:\ldbot\pix\sendpix1.png', grayscale=False, minSearchTime=6, confidence=.6)
            if ploc != None:
                pyautogui.click(ploc)
                break
            elif qloc != None:
                pyautogui.click(qloc)
                break
        time.sleep(3)
        
        if (z % (20)) == 0: ###--- esvaziamento de cache a cada 40 contatos
            subprocess.Popen(['ipconfig', '/flushdns'])  ### limpa o cache DNS
            subprocess.Popen([r'C:\Program Files\CCleaner\CCleaner64.exe', '/auto'])  ### executa o ccleaner para limpeza geral
            pymsgbox.alert('EXECUTANDO LIMPEZA GERAL. FLUSH DNS E CCLEANER...',timeout=500)  ### avisa sobre a limpeza geral periodica
            
        if (z % (30)) == 0: ###--- limpeza de cache e DNS a cada 60 contatos
            time.sleep(3)
            pyautogui.press('f2') ### alterna para SETTINGS
            time.sleep(2)
            pyautogui.click(x=700, y=200) ### clica na janela para alternar
            time.sleep(3)
            pyautogui.click(x=1045, y=725) ### clear cache
            time.sleep(3)
            pymsgbox.alert('CACHE DO ZAP ESVAZIADO!!!',timeout=2000)  ### avisa que o cache do zap fora esvaziado com sucesso
            pyautogui.press('f2') ### alterna para SETTINGS
            time.sleep(2)
            pyautogui.click(x=700, y=200) ### clica na janela para alternar

    # ==================== BLOCO 010 ====================#
    if (bloco % intervalo) == 1: ### valida o caso de o envio ser para apenas um contato no caso o total de contatos nao ser multiplo de 5 com resto 1
        pyautogui.click(x=65, y=75)
        time.sleep(3)      
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

    email = 'rmatos.tec@gmail.com'
    send_to_email = emailCc #################
    cc = emailCc ############################
    subject = 'DISPARAZAP: Maquina #' + str(maquina) + ' | ' + str(tempo)
    message = 'Segue o relatorio emitido pelo robo DISPARAZAP\r\ncom detalhes da MAQUINA #' + str(
        maquina) + '.\r\n\r\n    >> ' + str(titulo[int(maquina) - 1]) + '\r\n    >> MENSAGENS PROGRAMADAS: ' + str(
        len(celular)) + '\r\n    >> MENSAGENS ENVIADAS: ' + str(bloco) + '\r\n    >> RENDIMENTO: ' + (
                "{:.0%}".format((bloco) / len(celular))) + '\r\n\r\nAtt., Equipe Disparazap 2020.'
    password = 'juizdefora2@'

    msg = MIMEMultipart()

    msg['Subject'] = subject
    msg['From'] = u'Robo Disparazap <rmatos.tec@gmail.com>'
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
    server.sendmail(email, send_to_email + cc, text)
    server.quit()

#################### ENCERRA O SCRIPT E AS APLICACOES OS RESPECTIVOS PCDs ############################
pidbanfile = r'z:\ldbot\pid\ban.pid'
pidban = int(open(pidbanfile, 'r').read())    
try:
    os.kill(pidban, signal.SIGTERM)
except OSError:
    pymsgbox.alert('PCD BANIMENTO ENCERRADO!',timeout=1000)

print('Assistente PCD Banimento encerrado!')
pyautogui.press('esc')
pymsgbox.alert('TESTE REALIZADO COM SUCESSO!',timeout=4000)  ### encerra o envio das mensagens da campanha ativa
sys.exit(0)
