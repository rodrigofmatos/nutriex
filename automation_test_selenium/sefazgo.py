import time
import pymsgbox
import os
import shutil
import mailses
import datetime
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By

os.system('clear')
### variables declaration ###
cpf = "701.111.111-11"
password = "11111111"
ie = "11111111"
d = datetime.date.today()
periodo = datetime.datetime.strftime((d.replace(day=1) - datetime.timedelta(days=1)).replace(day=d.day),'%m/%Y')
tempo = 1
bucket = '/Users/macbok/Downloads'
url = "https://portal.sefaz.go.gov.br/portalsefaz-apps/auth/login-form"
title = "Portal de Aplicações - Login | Secretaria de Estado da Economia de Goiás"
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", bucket)
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
#############################

class main():
    def setup_method(self):
        self.driver = webdriver.Firefox(firefox_profile=profile)
        self.vars = {}

    def encerrar_tudo(self):
        self.driver.quit()

    def login(self):
        self.driver.get(url)
        p0 = self.driver.window_handles[0]
        self.driver.execute_script("username.value='" + cpf + "'")
        self.driver.find_element(By.ID, "password").send_keys(password)
        self.driver.find_element(By.ID, "btnAuthenticate").click()

    def malhafina(self):
        self.driver.find_element(By.XPATH, "//h3[contains(.,\'Malha Fina Estadual\')]").click()
        time.sleep(tempo)
        p1 = self.driver.window_handles[1]
        self.driver.close()
        self.driver.switch_to.window(p1)
        time.sleep(tempo)
        self.driver.get("https://sistemas.sefaz.go.gov.br/smw/web/listarCriticasMes/listarCriticasMesContribuinte.jsf?param=N")
        time.sleep(tempo)

    def relatorio(self):
        self.driver.execute_script("campoInicio.value='" + periodo + "'")
        self.driver.find_element(By.XPATH, "//*[@id='campoInscricao']").send_keys(ie)
        self.driver.find_element(By.XPATH, "//*[@id='j_idt47']").click()
        time.sleep(tempo)

    def download(self):
        try:  ### tenta encontrar o agrupador principal
            self.driver.find_element(By.ID, "dtArquivoProcessado:0:j_idt92").click()
        except Exception as e:
            print ("Error: ", e)
            pymsgbox.alert('Sem REGISTROS para o periodo de ' + periodo,timeout=1000)
            self.driver.quit()
            sys.exit(1)
        else:
            time.sleep(tempo)
            try:  ### tenta encontrar o agrupador das saidas
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:0:j_idt115").click()
            except Exception as e:
                print ("Error: ", e)
                pymsgbox.alert('Sem SAIDAS para o periodo de ' + periodo,timeout=1000)
            else:
                time.sleep(tempo)
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:0:expButton").click()
                time.sleep(tempo)
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:0:j_idt130").click()
                time.sleep(tempo)
                filename = max([os.path.join(bucket, f) for f in os.listdir(bucket)], key=os.path.getctime)
                shutil.move(filename,os.path.join(bucket,ie + "_saidas_" + periodo.replace('/','') + ".xls"))

            try:  ### tenta encontrar o agrupador das entradas
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:1:j_idt115").click()
            except Exception as e:
                print ("Error: ", e)
                pymsgbox.alert('Sem ENTRADAS para o periodo de ' + periodo,timeout=1000)
            else:
                time.sleep(tempo)
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:1:expButton").click()
                time.sleep(tempo)
                self.driver.find_element(By.ID, "dtArquivoProcessado:0:dtProcessamentos:1:j_idt130").click()
                time.sleep(tempo)
                filename = max([os.path.join(bucket, f) for f in os.listdir(bucket)], key=os.path.getctime)
                shutil.move(filename,os.path.join(bucket,ie + "_entradas_" + periodo.replace('/','') + ".xls"))
        valor = []
        valor = [el.text for el in self.driver.find_elements(By.ID,"dtArquivoProcessado")]
        apuracao = ' '.join(map(str, valor)).replace('\n',';').replace(';expand;','\n')\
                    .replace('Arquivos Processados;Inscrição;Mes/Ano;Processado;Entrega EFD;'\
                    'Retificador;Inscrição;Nº de Críticas;Valor;Status;1 Arquivo(s) encontrado(s)\n','')\
                    .replace(';Operação;Nº de Críticas;Valor','').replace(';Descrição do Tipo da Critica;Quantidade;Valor','')\
                    .replace(';' + ie + ';',';').replace('Proc. ','')
        f = open("apuracao.csv", "w")
        f.write(apuracao)
        f.close()

run = main()
run.setup_method()
run.login()
run.malhafina()
run.relatorio()
run.download()
mailses.enviamsg()
run.encerrar_tudo()
