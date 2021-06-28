import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime

from pytz import timezone
import pandas as pd

filein = u'F:\\python\\mbsnexo\\S1.csv'
fileout = u'F:\\python\\mbsnexo\\fatnegocio.csv'


df = pd.read_csv(filein, skiprows=[1], sep=';', header=0, error_bad_lines=False, dtype='unicode')
cols = df.columns[[16, 17, 18, 19, 20, 21]]
df[cols] = df[cols].apply(pd.to_numeric, errors='coerce', axis=0)

df1 = df.loc[df['NEGOCIO'] != '||||SEM_CLASSIFICACAO||||']
dt1 = df1.groupby(['COLIGADA','NEGOCIO']).agg({'FATLIQ': 'sum', 'CTB': 'sum'})
dt1 = dt1.sort_values(by=['COLIGADA','FATLIQ'], ascending=False)
dt1 = dt1.applymap("${0:,.2f}".format)

df2 = df.loc[(df['NEGOCIO'] != '||||SEM_CLASSIFICACAO||||') & (df['ANOMES'] == '2018/11')]
dt2 = df2.groupby(['COLIGADA','NEGOCIO']).agg({'FATLIQ': 'sum', 'CTB': 'sum'})
dt2 = dt2.sort_values(by=['COLIGADA','FATLIQ'], ascending=False)
dt2 = dt2.applymap("${0:,.2f}".format)

time = datetime.datetime.now(timezone('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')
email = 'rodrigo.matos@mw.far.br'
send_to_email = ['rodrigo.matos@mw.far.br']
cc = ['rodrigo.matos@mw.far.br']#, 'vanessa.sousa@avanse.com.br']
subject = 'Faturamento Consolidado 2018 | ' + str(time)
html = """\
<html>
  <head></head>
  <body>
    <p>Olá!<br><br>
       Segue o relatório de faturamento com dados atualizados em """ + str(time) \
       + """.<br><h4>FATURAMENTO CONSOLIDADO ANO: 2018</h4>""" + str(dt1.to_html().replace('<thead>','<thead style = "background-color: green" "text-align: center">').replace('<td>','<td style="text-align: right">').replace('.',';').replace(',','.').replace(';',',')) \
       + """.<br><h4>FATURAMENTO NOV/2018</h4>""" + str(dt2.to_html().replace('<thead>','<thead style = "background-color: yellow" "text-align: center">').replace('<td>','<td style="text-align: right">').replace('.',';').replace(',','.').replace(';',',')) \
       + """<br><br>Mais gráficos, relatórios e filtros avançados disponíveis na ferramenta gratuita QSense.<br>
       Acesse o link <a href="https://qlikcloud.com/">Portal Nutriex</a> e boa navegação.<br><br>
       Att., Rodrigo Matos.
    </p>
  </body>
</html>
"""
password = 'juizdefora'
msg = MIMEMultipart()
msg['Subject'] = subject
msg['From'] = email
msg['To'] = ', '.join(send_to_email)
msg['Cc'] = ', '.join(cc)
body = html
msg.attach(MIMEText(body, 'html'))
server = smtplib.SMTP('correio.mileniofarma.com.br', 587)
server.starttls()
server.login(email, password)
text = msg.as_string()
server.sendmail(email, send_to_email + cc, text)
server.quit()
