from selenium import webdriver
import time
import os
import glob
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
import datetime

now = datetime.datetime.now()
pathcoleta = 'F:\\python\\fretes\\coleta\\'

files = glob.glob(pathcoleta + '*.xls')
for f in files:
    os.remove(f)

files = glob.glob(pathcoleta + 'quality*.*')
for f in files:
    os.remove(f)

driver = webdriver.Chrome('F:\\python\\fretes\\quality\\chromedriver.exe')
driver.get("https://sistema.qualityentregas.com.br/sistema")

# ACESSO A PAGINA DE LOGIN
u = driver.find_element_by_css_selector("input[type='text']")
u.send_keys('vdm')
p = driver.find_element_by_css_selector("input[type='password']")
p.send_keys('1234')
q = driver.find_element_by_css_selector("input[type='submit']")
q.click()

# ELIMINAR ALERTA DE ERRO APOS LOGIN
time.sleep(2)
try:
    driver.switch_to.alert.accept()
except:
    pass

# ACESSO A PAGINA DE EXTRACAO DE RELATORIOS
x = driver.find_element_by_xpath("//*[@id='navmenu-h']/li[5]/ul/li[2]/a").get_attribute('href')
driver.get(x)

# PARAMETROS PARA EXTRACAO DOS RELATORIOS EM EXCEL
driver.find_element_by_name('FiDatainicial').send_keys(now.strftime('01/%m/%Y'))
driver.find_element_by_name('ChkXls').click()

# EXTRAI O RELATORIO PARTE 1 DE 3 PARA CAPTURAR TODAS AS COLUNAS
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_id('id').click()
driver.find_element_by_id('RC').click()
driver.find_element_by_id('RL').click()
driver.find_element_by_id('RT1').click()
driver.find_element_by_id('RT2').click()
driver.find_element_by_id('RT3').click()
driver.find_element_by_id('RT4').click()
driver.find_element_by_id('RT5').click()
driver.find_element_by_id('RT6').click()
driver.find_element_by_id('bairro').click()
driver.find_element_by_id('cep').click()
driver.find_element_by_id('chaveacesso').click()
driver.find_element_by_id('cidade').click()
driver.find_element_by_id('cliente').click()
driver.find_element_by_name('FbGerar').submit()

# EXTRAI O RELATORIO PARTE 2 DE 3 PARA CAPTURAR TODAS AS COLUNAS
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_id('coletas.prazo').click()
driver.find_element_by_id('ctecomp').click()
driver.find_element_by_id('datacadastro').click()
driver.find_element_by_id('dataretorno').click()
driver.find_element_by_id('dest_cpfcnpj').click()
driver.find_element_by_id('destinatario').click()
driver.find_element_by_id('diasatraso').click()
driver.find_element_by_id('diasemana').click()
driver.find_element_by_id('emissaocte').click()
driver.find_element_by_id('entrega').click()
driver.find_element_by_id('estado').click()
driver.find_element_by_id('frete').click()
driver.find_element_by_id('idpedidooriginal').click()
driver.find_element_by_id('idregiao').click()
driver.find_element_by_id('idregional').click()
driver.find_element_by_id('idusuariocadastrou').click()
driver.find_element_by_name('FbGerar').submit()

# EXTRAI O RELATORIO PARTE 3 DE 3 PARA CAPTURAR TODAS AS COLUNAS
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_name('FsTodos').click()
driver.find_element_by_id('jc').click()
driver.find_element_by_id('logradouro').click()
driver.find_element_by_id('motivo').click()
driver.find_element_by_id('numcte').click()
driver.find_element_by_id('numero').click()
driver.find_element_by_id('numerocontrole').click()
driver.find_element_by_id('origem').click()
driver.find_element_by_id('periodo').click()
driver.find_element_by_id('peso').click()
driver.find_element_by_id('recebidopor').click()
driver.find_element_by_id('rota').click()
driver.find_element_by_id('soentregar').click()
driver.find_element_by_id('status').click()
driver.find_element_by_id('tentativas').click()
driver.find_element_by_id('tipo').click()
driver.find_element_by_id('valor').click()
driver.find_element_by_id('valorcte').click()
driver.find_element_by_id('valorctecomp').click()
driver.find_element_by_id('volumes').click()
driver.find_element_by_name('FbGerar').submit()

# RENOMEAR OS RELATORIOS CAPTURADOS ANTES DA TRANSFORMACAO FINAL
arquivo = pathcoleta + 'quality_' + now.strftime('%d%m%y') + '.csv'
arquivo1 = pathcoleta + 'quality001_' + now.strftime('%d%m%y') + '.csv'
arquivo2 = pathcoleta + 'quality002_' + now.strftime('%d%m%y') + '.csv'
arquivo3 = pathcoleta + 'quality003_' + now.strftime('%d%m%y') + '.csv'

files = glob.glob(pathcoleta + 'relatorio.xls')
files1 = glob.glob(pathcoleta + 'relatorio.xls')
files2 = glob.glob(pathcoleta + 'relatorio (1).xls')
files3 = glob.glob(pathcoleta + 'relatorio (2).xls')

continua = True
while continua:
    if files1 == '[]' : break
    files1 = glob.glob(pathcoleta + 'relatorio.xls')
    time.sleep(0.1)
    for file in files1:
        os.rename(file, arquivo1)
        continua = False

continua = True
while continua:
    if files2 == '[]' : break
    files2 = glob.glob(pathcoleta + 'relatorio (1).xls')
    time.sleep(0.1)
    for file in files2:
        os.rename(file, arquivo2)
        continua = False

continua = True
while continua:
    if files3 == '[]' : break
    files3 = glob.glob(pathcoleta + 'relatorio (2).xls')
    time.sleep(0.1)
    for file in files3:
        os.rename(file, arquivo3)
        continua = False

# FECHA O SITE E ENCERRA TUDO PARA ENVIO DO EMAIL
time.sleep(2)
driver.quit()

with open(arquivo1, 'r') as file:
    for line1 in file:
        if (line1.__len__())>1000:
            with open(arquivo2, 'r') as file:
                for line2 in file:
                    if (line2.__len__()) > 1000:
                        with open(arquivo3, 'r') as file:
                            for line3 in file:
                                if (line3.__len__()) > 1000:
                                    line = line1.strip() + line2.strip() + line3.strip()
                                    line = line.replace("<tr>", "").replace("</td>", ";").replace("&amp; ","").replace("""<td  align="center">""", "").replace("<td nowrap >", "").replace("""<td align="center">""","").replace("""<type="text" style="cursor: pointer;" onClick="visualizaNotas""","").replace("<td  >","") \
    .replace("""<td  align="right">""","").replace("</tr>","|").replace(""""/>""","").replace("</tbody>","").replace("  "," ").replace("---","").replace("00/00/0000 00:00:00","01/01/1900 00:00:01").replace(";0,00;",";0;").upper().strip()
                                    line = line.split('|')
                                    i = 0
                                    j = 1
                                    f = line.__len__() // 3
                                    with open(arquivo, 'a') as file:
                                        file.writelines('CODE;CLIENTE;RC;DATA_RC;RT1;DATA_RT1;RT2;DATA_RT2;RT3;DATA_RT3;RT4;DATA_RT4;RT5;DATA_RT5;RT6;DATA_RT6;RL;DATA_RL;BAIRRO;CEP;CIDADE;ID;CHAVE_ACESSO;CODE;CONTROLE_PAI;DESTINATARIO;ENTREGA;CPF_CNPJ;FRETE;ESTADO;CDD;REGIAO;CTE_COMPLEMENTAR;DATA_CTE;DIA_ENTREGA;DATA_BAIXA;DATA_COLETA;USUARIO_CADASTRO;DATA_CADASTRO;DIAS_ATRASO;CODE;CONTROLE_ID;CONTROLE_FILHO;PERIODO;VALOR;JC;TIPO;STATUS;QL_NUM;PGTO;ORIGEM_CUB;MOTIVO;RECEBIDO_POR;VOLUMES;PESO_COBRADO;LOGRADOURO;CTE;ROTA;VALOR_CTE;VALOR_COMP;TENTATIVAS\n')
                                        while j <= f:
                                            while i < line.__len__():
                                                file.writelines(line[i])
                                                i = i + f
                                            i = j
                                            j = j + 1
                                            file.writelines('\n')

files = glob.glob(pathcoleta + 'quality0*.*')
for f in files:
    os.remove(f)

#ENVIA EMAIL DO RELATORIO FINAL PARA OS DESTINATARIOS DO PROJETO
email = 'rodrigo.matos@mw.far.br'
send_to_email = ['andrea.ferraz@mw.far.br']
cc = ['rodrigo.matos@mw.far.br','rafael.augusto@vdmoplog.com.br']
subject = 'Relatorio Diario de Fretes (Transportadora QUALITY/TRANSFARMA)'
message = 'Ola! Segue o relatorio extraido do portal da transportadora em formato CSV. Att., Rodrigo Matos.'
file_location = arquivo
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
