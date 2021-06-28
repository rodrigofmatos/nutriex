import pymsgbox
import pandas as pd         
import sys
import pyodbc
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
import time
import ntpath
import numpy as np
import xlsxwriter

start_time = time.time() ###-- inicia a temporizacao do codigo
now = datetime.datetime.now()
now_str = now.strftime("%Y%m%d-%Hh%Mm")
emailTo = ['sergio.rosa@nutriex.com.br']
emailCc = ['rodrigo.matos@nutriex.com.br']

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=TELAVIV\ERP;'
                      'Database=MBS;'
                      'UID=sa;'
                      'PWD=timilenio')

cursor = conn.cursor()

db_result= r'f:\bot\rpa010\rpa010_{}.xlsx'.format(now_str) ###-- arquivo final que sera enviado por email

df_sqlvenda = """
                SELECT
                    [Filial] + '_' + [TipoMaterial] EMPRESA
                    ,[Codigo] + '_' + [Descricao] ITEM
                    ,[DepositoCodigo] + '_' + [DepositoNome] DEPOSITO
                    ,ENDERECO     
                    ,LOTE
                    ,coalesce([Fabricacao],'1/1/1900') FABRICACAO
                    ,VALIDADE
                    ,CASE WHEN [DiasVencimento] > 0 THEN 'VENCIDO' WHEN [DiasVencimento] = 0 THEN 'VENCE HOJE' ELSE 'VENC. < 180DD' END STATUS
                        ,cast(SUM([Quantidade]) as float) QTD
                        ,cast(SUM([VlTotal]) as float) VLR
                FROM [DBMilenioWMS].[dbo].[Monitoramento_Estoque_Lotes]
                where [DepositoNome] NOT LIKE 'DIV%' AND lote is not null and validade is NOT null and [DiasVencimento] >= -180 and (((case when filial like 'Innovapharma  Lab.%' or filial like 'Nutriex%' then 'i' else 'd' end) = 'i' AND (tipomaterial like '%prima%' or tipomaterial like '%embalagem%')) or ((case when filial like 'Innovapharma  Lab.%' or filial like 'Nutriex%' then 'i' else 'd' end) = 'd' AND (tipomaterial like '%revenda%' or tipomaterial like '%acabado%')))
                GROUP BY [Filial] + '_' + [TipoMaterial],[Codigo] + '_' + [Descricao]   ,[DepositoCodigo] + '_' + [DepositoNome]   ,Lote ,[Fabricacao] ,[Endereco],Validade , CASE WHEN [DiasVencimento] > 0 THEN 'VENCIDO' WHEN [DiasVencimento] = 0 THEN 'VENCE HOJE' ELSE 'VENC. < 180DD' END, (case when filial like 'Innovapharma  Lab.%' or filial like 'Nutriex%' then 'i' else 'd' end)
            """

df_sqlqueryv = pd.read_sql_query(df_sqlvenda,conn,index_col=None)
df = df_sqlqueryv.apply(pd.Series)

workbook = xlsxwriter.Workbook(db_result)
worksheet = workbook.add_worksheet('result')
row = 1
col = 0

colunas = ['EMPRESA','ITEM','DEPOSITO','ENDERECO','LOTE','FABRICACAO','VALIDADE','STATUS','QTD','VLR']
worksheet.write_row(0,0,colunas,workbook.add_format({'bold': True, 'align': 'center'}))
worksheet.freeze_panes(1, 0)
worksheet.autofilter(0, 0, 0, 9)
worksheet.set_column(0, 7, 17)
worksheet.set_column(8, 9, 8)

while row-1 < len(df):
    if df['STATUS'][row-1] == 'VENCIDO':
        worksheet.write(row,0,df['EMPRESA'][row-1],workbook.add_format({'bold': False, 'font_color': 'red'}))
        worksheet.write(row,1,df['ITEM'][row-1],workbook.add_format({'bold': False, 'font_color': 'red'}))
        worksheet.write(row,2,df['DEPOSITO'][row-1],workbook.add_format({'bold': False, 'font_color': 'red'}))
        worksheet.write(row,3,df['ENDERECO'][row-1],workbook.add_format({'bold': False, 'font_color': 'red'}))
        worksheet.write(row,4,df['LOTE'][row-1],workbook.add_format({'bold': False, 'font_color': 'red'}))
        worksheet.write(row,5,df['FABRICACAO'][row-1],workbook.add_format({'align': 'center','bold': False, 'font_color': 'red'}))
        worksheet.write(row,6,df['VALIDADE'][row-1],workbook.add_format({'align': 'center','bold': False, 'font_color': 'red'}))
        worksheet.write(row,7,df['STATUS'][row-1],workbook.add_format({'bold': True, 'font_color': 'red'}))
        worksheet.write(row,8,(df['QTD'][row-1]),workbook.add_format({'num_format': 3,'bold': True, 'font_color': 'red'}))
        worksheet.write(row,9,(df['VLR'][row-1]),workbook.add_format({'num_format': 3, 'bold': True, 'font_color': 'red'}))
    elif df['STATUS'][row-1] == 'VENCE HOJE':
        worksheet.write(row,0,df['EMPRESA'][row-1],workbook.add_format({'bold': False, 'font_color': 'orange'}))
        worksheet.write(row,1,df['ITEM'][row-1],workbook.add_format({'bold': False, 'font_color': 'orange'}))
        worksheet.write(row,2,df['DEPOSITO'][row-1],workbook.add_format({'bold': False, 'font_color': 'orange'}))
        worksheet.write(row,3,df['ENDERECO'][row-1],workbook.add_format({'bold': False, 'font_color': 'orange'}))
        worksheet.write(row,4,df['LOTE'][row-1],workbook.add_format({'bold': False, 'font_color': 'orange'}))
        worksheet.write(row,5,df['FABRICACAO'][row-1],workbook.add_format({'align': 'center','bold': False, 'font_color': 'orange'}))
        worksheet.write(row,6,df['VALIDADE'][row-1],workbook.add_format({'align': 'center','bold': False, 'font_color': 'orange'}))
        worksheet.write(row,7,df['STATUS'][row-1],workbook.add_format({'bold': True, 'font_color': 'orange'}))
        worksheet.write(row,8,(df['QTD'][row-1]),workbook.add_format({'num_format': 3,'bold': True, 'font_color': 'orange'}))
        worksheet.write(row,9,(df['VLR'][row-1]),workbook.add_format({'num_format': 3, 'bold': True, 'font_color': 'orange'}))
    else:
        worksheet.write(row,0,df['EMPRESA'][row-1])
        worksheet.write(row,1,df['ITEM'][row-1])
        worksheet.write(row,2,df['DEPOSITO'][row-1])
        worksheet.write(row,3,df['ENDERECO'][row-1])
        worksheet.write(row,4,df['LOTE'][row-1])
        worksheet.write(row,5,df['FABRICACAO'][row-1],workbook.add_format({'align': 'center'}))
        worksheet.write(row,6,df['VALIDADE'][row-1],workbook.add_format({'align': 'center'}))
        worksheet.write(row,7,df['STATUS'][row-1])
        worksheet.write(row,8,(df['QTD'][row-1]),workbook.add_format({'num_format': 3}))
        worksheet.write(row,9,(df['VLR'][row-1]),workbook.add_format({'num_format': 3}))
    row = row + 1

workbook.close()

arquivo = db_result
tempo = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
email = 'robot@nutriex.com.br'
send_to_email = emailTo #################
cc = emailCc ############################
subject = "MONITORAMENTO DIÁRIO DE VALIDADES DO ESTOQUE | " + str(tempo)
message = "Segue o relatório diário contendo a lista de itens vencidos, vencendo hoje e próximos ao vencimento (180 dias).\r\n\r\n<Tempo total de processamento: %.1f segundos>" % (time.time() - start_time) + "\r\n\r\nAtt., Equipe Inteligência 2020."
password = '111111'
msg = MIMEMultipart()
msg['Subject'] = subject
msg['From'] = u'Robo Inteligencia <robot@nutriex.com.br>'
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
server = smtplib.SMTP('correio.mileniofarma.com.br', 587)
server.starttls()
server.login(email, password)
text = msg.as_string()
server.sendmail(email, send_to_email+cc, text)
server.quit()

sys.exit(0)