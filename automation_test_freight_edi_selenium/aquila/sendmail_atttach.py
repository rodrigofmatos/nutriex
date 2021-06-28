import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
import datetime

now = datetime.datetime.now()

email = 'rodrigo.matos@mw.far.br'
send_to_email = ['vipmenbrasil@gmail.com']
cc = ['rodrigo.matos@mw.far.br']
subject = 'Relatorio Diario de Fretes (Transportadora AQUILA)'
message = 'Ola! Segue o relatorio extraido do portal da transportadora em formato CSV. Att., Rodrigo Matos.'
file_location = 'C:\\Users\\Bulkleaders\\Downloads\\aquila_'+now.strftime('%d%m%y')+'.csv'
password = 'juizdefora'

msg = MIMEMultipart()

msg['Subject'] = subject
msg['From'] = email
msg['To'] = ', '.join(send_to_email)
msg['Cc'] = ', '.join(cc)

body = message

msg.attach(MIMEText(body, 'plain'))

filename = ntpath.basename(file_location)
attachment = open(file_location, "rb")

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
