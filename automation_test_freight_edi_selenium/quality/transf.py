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
import re

now = datetime.datetime.now()
pathcoleta = 'F:\\python\\fretes\\coleta\\'

arquivo = pathcoleta + 'quality_' + now.strftime('%d%m%y') + '.csv'
arquivo1 = pathcoleta + 'quality001_' + now.strftime('%d%m%y') + '.csv'
arquivo2 = pathcoleta + 'quality002_' + now.strftime('%d%m%y') + '.csv'
arquivo3 = pathcoleta + 'quality003_' + now.strftime('%d%m%y') + '.csv'

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
                                        file.writelines('CODE;CLIENTE;RC;DATA_RC;RT1;DATA_RT1;RT2;DATA_RT2;RT3;DATA_RT3;RT4;DATA_RT4;RT5;DATA_RT5;RT6;DATA_RT6;RL;DATA_RL;AUTORIZADO_POR;BAIRRO;CEP;CIDADE;ID;CHAVE_ACESSO;CODE;CONTROLE_PAI;DESTINATARIO;ENTREGA;CPF_CNPJ;FRETE;ESTADO;CDD;REGIAO;CTE_COMPLEMENTAR;DATA_CTE;DIA_ENTREGA;DATA_BAIXA;DATA_COLETA;USUARIO_CADASTRO;DATA_CADASTRO;DIAS_ATRASO;CODE;CONTROLE_ID;CONTROLE_FILHO;PERIODO;VALOR;JC;TIPO;STATUS;QL_NUM;PGTO;ORIGEM_CUB;MOTIVO;RECEBIDO_POR;VOLUMES;PESO_COBRADO;LOGRADOURO;CTE;ROTA;VALOR_CTE;VALOR_COMP;TENTATIVAS\n')
                                        while j <= f:
                                            while i < line.__len__():
                                                file.writelines(line[i])
                                                i = i + f
                                            i = j
                                            j = j + 1
                                            file.writelines('\n')