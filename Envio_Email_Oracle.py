from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders 
import schedule 
import time
import cx_Oracle
import smtplib
import csv

#DEFININDO TIMER PARA EXECUÇÃO AUTOMATICA DO SCRIPT
#def script():

#-------------------------------------------------------CONECTANDO COM A BANCO ORACLE-----------------------------------------
#CONECTANDO NO BANCO
con = cx_Oracle.connect("usuario/senha@ip-servidor/instancia")
cursor = con.cursor()

#GERANDO O ARQUIVO CSV 
csv_file = open("arquivo.csv", "w")

#ESTRUTURANDO O ARQUIVO E EXECUTANDO SQL
writer = csv.writer(csv_file, delimiter=';', lineterminator="\n", quoting=csv.QUOTE_NONNUMERIC)
r = "select * FROM"
cursor.execute(r)

#NOME DA TABELA
col_names = [row[0] for row in cursor.description]
writer.writerow(col_names)

#GRAVANDO INFORMAÇÕES DA TABELA
for row in cursor:
	writer.writerow(row)
   
cursor.close()
con.close()
csv_file.close()

print('Exportação Concluída !')

#----------------------------------------------------CONECTANDO COM SMTP E ENVIANDO E-MAIL-----------------------------------

msg = MIMEMultipart()

msg["Subject"] = "Example"
msg["From"] = "email-exemplo@hotmail.com"
msg["To"] = "email-exemplo@hotmail.com"
msg["Cc"] = "email-exemplo@hotmail.com"

#msg['From'] = email_user
#msg['To'] = email_send 
#msg['Cc'] = cc
#msg['Subject'] = subject

print ('Enviando E-mail...\n')

body = 'Teste E-mail'
msg.attach(MIMEText(body,'plain'))

filename='arquivo.csv'
attachment  = open(filename,'rb')

part = MIMEBase('application' , 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition' , "attachment; filename = " + filename)

msg.attach(part)
text = msg.as_string()
#UTILIZANDO SERVIDOR SEGURO 
server = smtplib.SMTP("smtp-mail.outlook.com",587)
#server = smtplib.SMTP("smtp-mail.outlook.com",587)
#server.starttls()
#server.login(email_user,email_password)

server.sendmail(msg["From"], msg["To"].split(",") + msg["Cc"].split(",") , text)

print ('E-mail enviado!')

server.quit()
