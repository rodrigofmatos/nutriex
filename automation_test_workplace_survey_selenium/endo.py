import random
from pywinauto.keyboard import send_keys
import datetime
import subprocess
import pyautogui
import time
import sys

pyautogui.FAILSAFE = False
i = 0
j = 0
k = 0
m = 0

site = 'https://docs.google.com/forms/d/e/1FAIpQLScAt2JsFME4VLHcTJuBg6I09h0Rr_GSlOeVK0rY4lHLD-YRzQ/viewform'
enviados = r'f:\python\rh\enviados.pid'

subprocess.Popen([r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', site]) 

while j < 20:
    time.sleep(3) ### define o tempo de carregamento inicial da pagina

    send_keys('{TAB}') #001. Gênero 
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,5)
    pyautogui.press('down',presses=i)
 
    send_keys('{TAB}') #002. Qual área da empresa você faz parte? 
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,15)
    pyautogui.press('down',presses=i)

    send_keys('{TAB}') #003. Há quanto tempo você trabalha nessa empresa?
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,15)
    pyautogui.press('down',presses=i)

    send_keys('{TAB}') #004. De 1 a 10, quanto você se sente orgulhoso de fazer parte da empresa?
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,15)
    pyautogui.press('right',presses=i)

    send_keys('{TAB}') #005. Qual ação/atividade realizada em 2020 que você mais gostou?
    time.sleep(0.5)
    i = random.randint(0,10)
    pyautogui.press('TAB',presses=i)
    send_keys('{SPACE}')
    time.sleep(0.5)
    m = 12-i
    pyautogui.press('TAB',presses=m)
    time.sleep(0.5)

    send_keys('{TAB}') #000. Qual ação/atividade realizada em 2020 que você menos gostou? Por que?
    i = random.randint(0,30)
    if i == 0:
            pyautogui.typewrite('Mercadinho Peg Pag')
    elif i == 1:
        pyautogui.typewrite('Vulnerabilidade')
    elif i == 2:
        pyautogui.typewrite('Pote da Gratidao')
    elif i == 3:
        pyautogui.typewrite(' ')
    elif i == 4:
        pyautogui.typewrite('Dia das Mulheres')
    elif i == 5:
        pyautogui.typewrite('Dia das Maes')
    elif i == 6:
        pyautogui.typewrite(' ')
    elif i == 7:
        pyautogui.typewrite('Dia do Amigo')
    elif i == 8:
        pyautogui.typewrite('Dia dos Pais')
    elif i == 9:
        pyautogui.typewrite(' ')
    elif i == 10:
        pyautogui.typewrite('Setembro Amarelo')
    elif i == 11:
        pyautogui.typewrite('Outubro Rosa')
    elif i == 12:
        pyautogui.typewrite('Dia da Gentileza')
    elif i == 13:
        pyautogui.typewrite(' ')
    elif i == 14:
        pyautogui.typewrite('Novembro Azul (Go Fit Truck)')
    else:
        pyautogui.typewrite(' ')


    send_keys('{TAB}') #000. Que tipos de ações/atividades você gostaria que fossem feitas na empresa?
    time.sleep(0.5)
    
    send_keys('{TAB}') #001. Dentro dos benefícios da empresa, quais você mais gosta e acha importante?
    time.sleep(0.5)
    i = random.randint(0,8)
    pyautogui.press('TAB',presses=i)
    send_keys('{SPACE}')
    time.sleep(0.5)
    m = 10-i
    pyautogui.press('TAB',presses=m)
    time.sleep(0.5)

    send_keys('{TAB}') #000. Quais benefícios você gostaria que tivesse na empresa?
    i = random.randint(0,40)
    if i == 0:
        pyautogui.typewrite('COBRIR O ESTACIONAMENTO EXTERNO PARA CARROS E MOTOS')
    elif i == 1:
        pyautogui.typewrite('aulas de ingles de graca')
    elif i == 2:
        pyautogui.typewrite(' ')
    elif i == 3:
        pyautogui.typewrite('curso de Espanhol gratuito')
    elif i == 4:
        pyautogui.typewrite('sorteio de viagem pra disnei')
    elif i == 5:
        pyautogui.typewrite(' ')
    elif i == 6:
        pyautogui.typewrite('creche')
    elif i == 7:
        pyautogui.typewrite('home office')
    elif i == 8:
        pyautogui.typewrite('Faculdade corporativa.')
    elif i == 9:
        pyautogui.typewrite('Telefone 24 horas para pedir conselhos')
    elif i == 10:
        pyautogui.typewrite('dia da gratidao')
    elif i == 11:
        pyautogui.typewrite('cargos e salarios')
    elif i == 12:
        pyautogui.typewrite(' ')
    elif i == 13:
        pyautogui.typewrite('flexibilida de horario')
    elif i == 14:
        pyautogui.typewrite('informar sempre')
    elif i == 15:
        pyautogui.typewrite('curso de tempo')
    elif i == 16:
        pyautogui.typewrite('feedback')
    elif i == 17:
        pyautogui.typewrite(' ')
    elif i == 18:
        pyautogui.typewrite('metas definidas')
    elif i == 19:
        pyautogui.typewrite('um chefe somente')
    elif i == 20:
        pyautogui.typewrite('pedir uma coisa de cada vez')
    elif i == 21:
        pyautogui.typewrite('cafeteira gratis')
    elif i == 22:
        pyautogui.typewrite(' ')
    elif i == 23:
        pyautogui.typewrite('2h almoco')
    elif i == 24:
        pyautogui.typewrite('rede pra dormir')
    elif i == 25:
        pyautogui.typewrite('desafios')
    elif i == 26:
        pyautogui.typewrite(' ')
    elif i == 27:
        pyautogui.typewrite('trazer familia pra conhecer fabrica')
    elif i == 28:
        pyautogui.typewrite(' ')
    elif i == 29:
        pyautogui.typewrite('ginastica')
    else:
        pyautogui.typewrite(' ')
        
    send_keys('{TAB}') #001. Com que frequência você lê os comunicados internos (enviados por e-mail, TV corporativa e Spark)?
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,15)
    pyautogui.press('right',presses=i)

    send_keys('{TAB}') #001. Qual o canal que você mais acessa para ler os comunicados? 
    time.sleep(0.5)
    send_keys('{SPACE}')
    i = random.randint(0,8)
    pyautogui.press('down',presses=i)


    send_keys('{TAB}') #001. SUBMIT
    send_keys('{ENTER}') #001. CLICAR BOTAO ENVIO
    time.sleep(3)

    send_keys('{TAB}') #001. nova resposta
    send_keys('{ENTER}') #001. confirmar nova resposta e recarregar o formulario

    open(enviados, 'a').write('registro: ' + str(j) + ' as ' + str(datetime.datetime.now()) + '\n')
    j = j + 1

    time.sleep(3)

# send_keys('%{F4}')
# time.sleep(0.5)
# send_keys('{ENTER}')
sys.exit(0)
