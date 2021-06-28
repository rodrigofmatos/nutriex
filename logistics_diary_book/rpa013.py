import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from oauth2client.service_account import ServiceAccountCredentials
import time
import sys
import pyodbc
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime

now = datetime.datetime.now().strftime("%H")
if int(now) >= 7 and int(now) <= 22:
    start_time = time.time() ###-- inicia a temporizacao do codigo
    emailTo = ['rodrigo.matos@nutriex.com.br']
    emailCc = ['rodrigo.matos@nutriex.com.br']

    #### FAZ A CONEXAO COM A PLANILHA GOOGLE SHEETS
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('nutriex-54f72803516b.json', scope)
    gc = gspread.authorize(credentials)
    spreadsheet_key = '1v1GXhUdFVATd-r7zlQJuHNasUDh0EyO1cO3FUsCz4-U'

    conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=srv2-sap\erp;'
                        'Database=MBS;'
                        'UID=sa;'
                        'PWD=@SapN22RwDt@')

    cursor = conn.cursor()

    df_sqlbook = """
                    SELECT [ORIGEM]
                        ,[EMPRESA]
                        ,[USO]
                        ,[ITEMCODE]
                        ,[ITEMNAME]
                        ,[QTDE]
                        ,[CAIXA]
                        ,[PESO]
                        ,[UNIDADE]
                        ,[SUPERVISOR]
                        ,[DTNF]
                        ,[FATLIQ]
                        ,[VERSION]
                    FROM [Mbs].[dbo].[rpa013]
                """

    ###-- LIMPA A PLANILHA ANTES DE COLAR OS NOVOS DADOS
    spreadsheet = gc.open("booklog")
    wks = spreadsheet.worksheet('FATMENSAL').clear()

    #### GRAVA AS INFORMACOES DO ARQUIVO CSV NA PLANILHA DO GOOGLE SHEEETS
    df = pd.read_sql_query(df_sqlbook,conn,index_col=None)
    wks = d2g.upload(df, spreadsheet_key, wks_name='FATMENSAL', credentials=credentials, row_names=False)

    tempo = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
    email = 'robot@nutriex.com.br'
    send_to_email = emailTo #################
    cc = emailCc ############################
    subject = 'BOOK LOGISTICA GERADO COM SUCESSO | ' + str(tempo)
    message = "Segue o BOOK DIARIO gerado.\r\n\r\n<Tempo total de processamento: %.1f segundos>" % (time.time() - start_time) + "\r\n\r\nAtt., Equipe TECNOLOGIA 2020."
    password = '111111'
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = u'Robo Tecnologia <robot@nutriex.com.br>'
    msg['To'] = ', '.join(send_to_email)
    msg['Cc'] = ', '.join(cc)
    body = message
    msg.attach(MIMEText(body, 'plain'))
    server = smtplib.SMTP('correio.mileniofarma.com.br', 587)
    server.starttls()
    server.login(email, password)
    text = msg.as_string()
    server.sendmail(email, send_to_email+cc, text)
    server.quit()

sys.exit(0)