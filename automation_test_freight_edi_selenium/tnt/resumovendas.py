
time = datetime.datetime.now(timezone('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')
email = 'rodrigo.matos@mw.far.br'
send_to_email = ['rodrigo.matos@mw.far.br']
cc = ['rodrigo.matos@mw.far.br']#, 'vanessa.sousa@avanse.com.br']
subject = 'Geracao do Relatorio de Faturamento Real - Nutriex | ' + str(time)
#message = 'Segue o relatorio de faturamento com dados de ' + str(time) + '\n\n' + str(dt) + '\n\n Att., Rodrigo Matos.'
html = """\
<html>
  <head></head>
  <body>
    <p>Olá!<br><br>
       Segue o relatório de faturamento com dados atualizados em """ + str(time) + """.<br><br> """ + str(dt.to_html()) + """
       Mais gráficos, relatórios e filtros avançados estão disponíveis na ferramenta gratuita QSense, <br>
       acessando o link do <a href="https://qlikcloud.com/">Portal Nutriex</a>.<br><br>
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
