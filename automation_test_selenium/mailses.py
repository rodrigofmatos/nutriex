import smtplib  
import os
import sys
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def enviamsg():
    os.system('clear')
    file = "apuracao.csv"
    headtitulo = ['ietit','pertit','dtprotit','dtefdtit','qtdtit','vlrtit','statustit']
    headopersai = ['operheadsai','qtdheadsai','vlrheadsai']
    headoperent = ['operheadent','qtdheadent','vlrheadent']
    headvalorent = ['critent','qtdcritent','vlrcritent']
    headvalorsai = ['critsai','qtdcritsai','vlrcritsai']
    lista = []
    f=open(file,"r")
    lines=len(open(file,"r").readlines())
    entrada = 0
    saida = 0

    if lines < 2:
            sys.exit(1)
    else:
            for i in range(lines):
                    line = f.readline().replace('\n','')
                    if i+1 == 1:
                            key1 = dict(zip(headtitulo,line.split(';')))
                            lista.append(key1)
                    elif line.find('Saída') > 0:
                            key2 = dict(zip(headopersai,line.split(';')))
                            lista.append(key2)
                            saida = i+1
                    elif line.find('Entrada') > 0:
                            key3 = dict(zip(headoperent,line.split(';')))
                            lista.append(key3)
                            entrada = i+1
                    else:
                            if entrada == 0:
                                    key4 = dict(zip(headvalorsai,line.split(';')))
                            else:
                                    key4 = dict(zip(headvalorent,line.split(';')))
                            lista.append(key4)
            f.close()

    ietit = ''.join([d['ietit'] for d in lista if 'ietit' in d])
    pertit = ''.join([d['pertit'] for d in lista if 'pertit' in d])
    dtprotit = ''.join([d['dtprotit'] for d in lista if 'dtprotit' in d])
    dtefdtit = ''.join([d['dtefdtit'] for d in lista if 'dtefdtit' in d])
    qtdtit = ''.join([d['qtdtit'] for d in lista if 'qtdtit' in d])
    vlrtit = ''.join([d['vlrtit'] for d in lista if 'vlrtit' in d])
    statustit = ''.join([d['statustit'] for d in lista if 'statustit' in d])

    operheadsai = ''.join([d['operheadsai'] for d in lista if 'operheadsai' in d])
    qtdheadsai = ''.join([d['qtdheadsai'] for d in lista if 'qtdheadsai' in d])
    vlrheadsai = ''.join([d['vlrheadsai'] for d in lista if 'vlrheadsai' in d])

    operheadent = ''.join([d['operheadent'] for d in lista if 'operheadent' in d])
    qtdheadent = ''.join([d['qtdheadent'] for d in lista if 'qtdheadent' in d])
    vlrheadent = ''.join([d['vlrheadent'] for d in lista if 'vlrheadent' in d])

    totalsai = len([d['critsai'] for d in lista if 'critsai' in d])
    totalent = len([d['critent'] for d in lista if 'critent' in d])
    totalsailista = []
    totalentlista = []
    total = []
    contador = 0

    if totalsai > totalent:
            contador = totalsai
    else:
            contador = totalent

    for i in range(totalsai):
            critsai = ''.join("""<td style="text-align: center;">""" + [d['critsai'] for d in lista if 'critsai' in d][i] + """</td>""")
            qtdcritsai = ''.join("""<td style="text-align: center;">""" + [d['qtdcritsai'] for d in lista if 'qtdcritsai' in d][i] + """</td>""")
            vlrcritsai = ''.join("""<td style="text-align: right;">""" + [d['vlrcritsai'] for d in lista if 'vlrcritsai' in d][i] + """</td>""")
            totalsailista.append("<tr>"+ critsai + qtdcritsai + vlrcritsai)

    if totalsai < contador:
            for j in range(contador-totalsai):
                    totalsailista.append("""<tr><td style="text-align: center;"> </td><td style="text-align: center;"> </td><td style="text-align: right;"> </td>""")

    for i in range(totalent):
            critent = ''.join("""<td style="text-align: center;">""" + [d['critent'] for d in lista if 'critent' in d][i] + """</td>""")
            qtdcritent = ''.join("""<td style="text-align: center;">""" + [d['qtdcritent'] for d in lista if 'qtdcritent' in d][i] + """</td>""")
            vlrcritent = ''.join("""<td style="text-align: right;">""" + [d['vlrcritent'] for d in lista if 'vlrcritent' in d][i] + """</td>""")
            totalentlista.append(critent + qtdcritent + vlrcritent + "</tr>")

    if totalent < contador:
            for j in range(contador-totalent):
                    totalentlista.append("""<td style="text-align: center;"> </td><td style="text-align: center;"> </td><td style="text-align: right;"> </td></tr>""")

    for i in range(contador):
            total.append(totalsailista[i] + totalentlista[i])

    criticas = ''.join(total)

    SENDER = 'rodrigo@email.com'  
    SENDERNAME = 'RODRIGO BOT'
    RECIPIENT  = "rodrigomatos@email.com"
    USERNAME_SMTP = "xxxxxxxxxxxxx"
    PASSWORD_SMTP = "x+xxxxxxxxxxxx"

    HOST = "email-smtp.us-east-1.amazonaws.com"
    PORT = 587

    f = open("apuracao.csv", "r")
    corpotexto = f.read()
    f.close()
    
    SUBJECT = 'Estatísticas do Portal Malha Fina - SEFAZ GO'

    BODY_HTML = """<!DOCTYPE html><html><head>
                    <style>
                    table {
                    border-collapse: separate;
                    border-spacing: 0;
                    color: #4a4a4d;
                    font: 14px/1.4 "Helvetica Neue", Helvetica, Arial, sans-serif;
                    }
                    th,
                    td {
                    padding: 5px 10px;
                    vertical-align: middle;
                    }
                    caption {
                    padding: 10px 15px;
                    vertical-align: middle;
                    background: #395870;
                    background: linear-gradient(#C3CDD5, #293f50);
                    color: #fff;
                    font-size: 20px;
                    font-weight: bold;
                    text-transform: uppercase;
                    }

                    thead {
                    background: #395870;
                    background: linear-gradient(#49708f, #293f50);
                    color: #fff;
                    font-size: 11px;
                    text-transform: uppercase;
                    }
                    th:first-child {
                    border-top-left-radius: 3px;
                    text-align: center;
                    }
                    th:last-child {
                    border-top-right-radius: 3px;
                    }
                    tbody tr:nth-child(even) {
                    background: #f0f0f2;
                    }
                    td {
                    border-bottom: 1px solid #cecfd5;
                    border-right: 1px solid #cecfd5;
                    }
                    td:first-child {
                    border-left: 1px solid #cecfd5;
                    }

                    tfoot {
                    text-align: right;
                    }
                    tfoot tr:last-child {
                    background: #f0f0f2;
                    color: #395870;
                    font-weight: bold;
                    }
                    tfoot tr:last-child td:first-child {
                    border-bottom-left-radius: 2px;
                    }
                    tfoot tr:last-child td:last-child {
                    border-bottom-right-radius: 2px;
                    }
                    
                    </style>

                    <title>Relatorio de criticas SEFAZ GO</title>
                    </head><body><br>
                    <table align="center" cellpadding="5" width="90%" style="font-family:Arial, Helvetica, sans-serif;font-size:80%">
                        <caption>RELATÓRIO DE CRÍTICAS SEFAZ GO</caption>
                        <thead>
                            <tr> 
                            <th>INSC. EST.</th>
                            <th>PERIODO</th>
                            <th>PROCESSADO</th>
                            <th>ENTREGA EFD</th>
                            <th>QTD</th>
                            <th>VALOR</th>
                            <th>STATUS</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="text-align: center;">""" + ietit + """</td>
                                <td style="text-align: center;">""" + pertit + """</td>
                                <td style="text-align: center;">""" + dtprotit + """</td>
                                <td style="text-align: center;">""" + dtefdtit + """</td>
                                <td style="text-align: center;">""" + qtdtit + """</td>
                                <td style="text-align: center;">""" + vlrtit + """</td>
                                <td style="text-align: center;">""" + statustit + """</td>
                            </tr>
                        </tbody>
                        <tbody><tr><td colspan="2" style="border:0px; padding-bottom:2em;"></td></tr></tbody>
                    </table>

                    <table align="center" cellpadding="5" width="90%" style="font-family:Arial, Helvetica, sans-serif;font-size:80%">
                        <thead>
                            <tr> 
                            <th>OPERACAO</th>
                            <th>QTD</th>
                            <th>VALOR</th>
                            <th>OPERACAO</th>
                            <th>QTD</th>
                            <th>VALOR</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr style="border-bottom: 3px solid #333335;">
                                <td style="text-align: center;"><strong>""" + operheadsai + """</strong></td>
                                <td style="text-align: center;"><strong>""" + qtdheadsai + """</strong></td>
                                <td style="text-align: right;"><strong>""" + vlrheadsai + """</strong></td>
                                <td style="text-align: center;"><strong>""" + operheadent + """</strong></td>
                                <td style="text-align: center;"><strong>""" + qtdheadent + """</strong></td>
                                <td style="text-align: right;"><strong>""" + vlrheadent + """</strong></td>
                            </tr>
                        </tbody>
                        <tbody>
                        """ + criticas + """
                        </tbody>
                        <tbody><tr><td colspan="2" style="border:0px; padding-bottom:2em;"></td></tr></tbody>
                    </table>
                    <p align="center" style="font-family:Arial, Helvetica, sans-serif;font-size:80%;font-weight:bold">@copyright 2021 INFOMACH</p></body></html>"""

    msg = MIMEMultipart('alternative')
    msg['Subject'] = SUBJECT
    msg['From'] = email.utils.formataddr((SENDERNAME, SENDER))
    msg['To'] = RECIPIENT
    part2 = MIMEText(BODY_HTML, 'html')
    msg.attach(part2)

    try:  
        server = smtplib.SMTP(HOST, PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(USERNAME_SMTP, PASSWORD_SMTP)
        server.sendmail(SENDER, RECIPIENT, msg.as_string())
        server.close()
    except Exception as e:
        print ("Erro: ", e)
    else:
        print ("Mensagem enviada!")
