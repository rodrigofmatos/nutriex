import sys
import urllib
import mysql.connector
import pandas as pd
import pyodbc
import sqlalchemy
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
from pytz import timezone

now = datetime.datetime.now().strftime("%H")
if int(now) >= 7 and int(now) <= 20:
  start_time = time.time()
  # emailTo = ['luciano.candido@nutriex.com.br']
  # emailCc = ['thamilla.araujo@nutriex.com.br']
  emailTo = ['rmatos.tec@gmail.com']
  emailCc = ['rmatos.tec@gmail.com']
  mydb = mysql.connector.connect(host="nutriex2.db.artia.com", user="cliente-95791", password="9I6vcDnKifOWLT3lIm014IyQ3VA=", database="artia")
  # engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect={}".format(urllib.parse.quote_plus("DRIVER=SQL Server Native Client 10.0;SERVER={0};PORT=1433;DATABASE={1};UID={2};PWD={3};TDS_Version=8.0;".format('192.168.1.74\erp', 'dbartia', 'sa', '@SapN22RwDt@'))))

  df_activities = pd.read_sql_query ("SELECT * FROM organization_95791_activities",mydb)
  df_community_users = pd.read_sql_query ("SELECT * FROM organization_95791_community_users",mydb)
  df_finances = pd.read_sql_query ("SELECT * FROM organization_95791_finances",mydb)
  df_folders = pd.read_sql_query ("SELECT * FROM organization_95791_folders",mydb)
  df_milestones = pd.read_sql_query ("SELECT * FROM organization_95791_milestones",mydb)
  df_organization_users = pd.read_sql_query ("SELECT * FROM organization_95791_organization_users",mydb)
  df_projects = pd.read_sql_query ("SELECT * FROM organization_95791_projects",mydb)
  df_teams_users = pd.read_sql_query ("SELECT * FROM organization_95791_teams_users",mydb)
  df_time_entries_new = pd.read_sql_query ("SELECT * FROM organization_95791_time_entries_new",mydb)
  df_users_skills = pd.read_sql_query ("SELECT * FROM organization_95791_users_skills",mydb)

  # df_activities.to_sql('activities', con=engine, if_exists='replace', index = False)
  # df_community_users.to_sql('community_users', con=engine, if_exists='replace', index = False)
  # df_finances.to_sql('finances', con=engine, if_exists='replace', index = False)
  # df_folders.to_sql('folders', con=engine, if_exists='replace', index = False)
  # df_milestones.to_sql('milestones', con=engine, if_exists='replace', index = False)
  # df_organization_users.to_sql('organization_users', con=engine, if_exists='replace', index = False)
  # df_projects.to_sql('projects', con=engine, if_exists='replace', index = False)
  # df_teams_users.to_sql('teams_users', con=engine, if_exists='replace', index = False)
  # df_time_entries_new.to_sql('time_entries_new', con=engine, if_exists='replace', index = False)
  # df_users_skills.to_sql('users_skills', con=engine, if_exists='replace', index = False)

  tempo = datetime.datetime.now(timezone('America/Fortaleza')).strftime('%d/%m/%Y %H:%M:%S')
  email = 'robot@nutriex.com.br'
  send_to_email = emailTo #################
  cc = emailCc ############################
  subject = 'DADOS DA NUVEM ARTIA SINCRONIZADOS COM SUCESSO | ' + str(tempo)
  message = "Os dados registrados nos bancos de dados do ARTIA foram baixados e enviados para o servidor da Nutriex.\r\n\r\n<Tempo total de processamento: %.1f segundos>" % (time.time() - start_time) + "\r\n\r\nAtt., Equipe TECNOLOGIA 2021."
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


  df_activities.to_excel(r'artia_atv.xlsx', sheet_name='atividades', engine='xlsxwriter', index = None, header=True)
  df_community_users.to_excel(r'artia_com.xlsx', sheet_name='comunidade', engine='xlsxwriter', index = None, header=True)
  df_finances.to_excel(r'artia_fin.xlsx', sheet_name='financeiro', engine='xlsxwriter', index = None, header=True)
  df_folders.to_excel(r'artia_pst.xlsx', sheet_name='pastas', engine='xlsxwriter', index = None, header=True)
  df_milestones.to_excel(r'artia_ctr.xlsx', sheet_name='controles', engine='xlsxwriter', index = None, header=True)
  df_organization_users.to_excel(r'artia_usr.xlsx', sheet_name='usuarios', engine='xlsxwriter', index = None, header=True)
  df_projects.to_excel(r'artia_prj.xlsx', sheet_name='projetos', engine='xlsxwriter', index = None, header=True)
  df_teams_users.to_excel(r'artia_eqp.xlsx', sheet_name='equipes', engine='xlsxwriter', index = None, header=True)
  df_time_entries_new.to_excel(r'artia_tmp.xlsx', sheet_name='tempo', engine='xlsxwriter', index = None, header=True)
  df_users_skills.to_excel(r'artia_hab.xlsx', sheet_name='habilidades', engine='xlsxwriter', index = None, header=True)
  sys.exit(0)
