import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from oauth2client.service_account import ServiceAccountCredentials
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

pathcoleta = u'f:\\bot\\rpaStudio\\'
files = pathcoleta + u'rpastudio.csv'

#### FAZ A CONEXAO COM A PLANILHA GOOGLE SHEETS
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('nutriex-54f72803516b.json', scope)
gc = gspread.authorize(credentials)
spreadsheet_key = '117sruFHwDjbJsp0yPQullL6QDd27CTx1VKBRBv9Lz7g'

#### GRAVA AS INFORMACOES DO ARQUIVO CSV NA PLANILHA DO GOOGLE SHEEETS
df = pd.read_csv(files, skiprows=[1], sep=',', header=0, error_bad_lines=False, dtype='unicode', encoding='utf-8', index_col=False)
wks = d2g.upload(df, spreadsheet_key, wks_name='Master', credentials=credentials, row_names=False)
time.sleep(10) #### aguarda 10 segundos para testar

#### ABRE A PLANILHA E CONTA OS REGISTROS GRAVADOS PARA VER SE A GRAVACAO FOI EFETUADA COM SUCESSO
spreadsheet = gc.open("nutriex1")
wks = spreadsheet.worksheet('Master')
# print('Linhas: ' + len(wks.col_values(1)).__str__() + ' | Colunas: ' + len(wks.row_values(1)).__str__() + ' | ' + wks.cell(2,21).value)

i = 0
while i < 3:
    if len(wks.row_values(1)) != 21 or len(wks.col_values(1)) < 500:
        #### ENVIA EMAIL QUE O PROCESSO FALHOU
        email = 'robot@nutriex.com.br'
        send_to_email = ['rodrigo.matos@nutriex.com.br']
        cc = ['rodrigo.matos@mw.far.br']
        subject = 'TENTATIVA #0' + str(i+1) + ': PROBLEMAS NA GERACAO DO RELATORIO DE TITULOS EM ABERTO'
        message = ('Linhas: ' + len(wks.col_values(1)).__str__() + ' | Colunas: ' + len(wks.row_values(1)).__str__() + ' | ' + wks.cell(2,21).value)
        password = '111111'

        msg = MIMEMultipart()

        msg['Subject'] = subject
        msg['From'] = u'Robo Inteligencia <robot@nutriex.com.br>'
        msg['To'] = ', '.join(send_to_email)
        msg['Cc'] = ', '.join(cc)

        body = message

        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('correio.mileniofarma.com.br', 587)
        server.starttls()
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, send_to_email + cc, text)
        server.quit()

        time.sleep(120) ## aguarda mais 2 minutos e tenta carregar novamente a planilha do google
        df = pd.read_csv(files, skiprows=[1], sep=',', header=0, error_bad_lines=False, dtype='unicode', encoding='utf-8', index_col=False)
        wks = d2g.upload(df, spreadsheet_key, wks_name='Master', credentials=credentials, row_names=False)
        time.sleep(10)  #### aguarda 10 segundos para testar
    else:
        break
    i = i + 1
