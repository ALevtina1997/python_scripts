# -*- coding: utf-8 -*-
import smtplib
import xlsxwriter
import MySQLdb
import datetime
import os
import sys

from smtplib import SMTP_SSL
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders
from email.mime.text import MIMEText

from datetime import timedelta
db = MySQLdb.connect('localhost','username','password','name_database')
#db.autocommit(True)
db.set_character_set('utf8')
cur = db.cursor()

cur.execute(""" CREATE TEMPORARY TABLE IF NOT EXISTS juridical_inet AS 
		( SELECT users.full_name, accounts.id 
		  FROM accounts,users,users_accounts,service_links,services_data  
		  WHERE users.is_juridical =1 
		  AND users.id = users_accounts.uid 
                  AND accounts.id = users_accounts.account_id 
                  AND accounts.block_id=0 
                  AND service_links.is_deleted = 0
                  AND service_links.account_id = accounts.id 
                  AND service_links.id > 0 
                  AND services_data.parent_service_id = 4 
                  AND services_data.id = service_links.service_id)""")

db.commit()

today=(datetime.date.today()).strftime("%Y-%m-%d")
last_week=(datetime.date.today() - datetime.timedelta(days=7)).strftime("%Y-%m-%d") #неделю назад

cur.execute("""SELECT juridical_inet.full_name, IFNULL(trafik.Mb, 0) AS traf  
	     FROM juridical_inet  
	     LEFT JOIN (SELECT account_id, SUM(bytes)/1024/1024 AS Mb  
			FROM discount_transactions_iptraffic_all
			WHERE discount_date>=UNIX_TIMESTAMP('2019-09-01 00:00:00' ) 
			AND discount_date<=UNIX_TIMESTAMP('""" +   today   + """') 
			AND t_class='10' GROUP BY account_id ) AS trafik 
			ON juridical_inet.id=trafik.account_id 
			 WHERE IFNULL(trafik.Mb, 0)<=50; """)

inet=cur.fetchall()
null_inet=[]
for i in inet:
	null_inet.append(i)

workbook = xlsxwriter.Workbook('/home/alevtina/juridical.xlsx')
worksheet = workbook.add_worksheet()

for j in range(len(null_inet)):
        worksheet.write(j, 0, str(null_inet[j][0]).decode('utf-8'))
workbook.close()

text = """Во вложении прикреплен файл с юр.лицами, у которых за предыдущую неделю входящий трафик менее 50Мбайт.\n Также ниже представлен список юр.лиц, у которых есть услуга интернет, но стоит системная блокировка, соответственно, трафика у них не будет. \n"""
cur.execute(""" SELECT users.full_name
	       FROM accounts,users,users_accounts,service_links,services_data, blocks_info
	       WHERE users.is_juridical =1 
	       AND users.id = users_accounts.uid 
	       AND accounts.id = users_accounts.account_id 
	       AND accounts.block_id= blocks_info.id
	       AND blocks_info.block_type =1
	       AND blocks_info.is_deleted=0
	       AND service_links.is_deleted = 0 
	       AND service_links.account_id = accounts.id 
	       AND service_links.id > 0 
	       AND services_data.parent_service_id = 4 
	       AND services_data.id = service_links.service_id """)

sistem=cur.fetchall()
system_block=[]
for i in sistem:
	system_block.append(i)

for n in range(len(system_block)):
	jurik=system_block[n][0]
	text=text + jurik + "\n"
	#print(type(system_block[n][0]))


filepath = "/home/alevtina/juridical.xlsx"
basename = os.path.basename(filepath)
address_to = ['1@mail.ru', '2@mail.ru']
address_from="server@company.ru"
# Compose attachment
part = MIMEBase('application', "octet-stream")
part.set_payload(open(filepath,"rb").read() )
Encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="%s"' % basename)

# Compose message
msg = MIMEMultipart()
msg['From'] = address_from
msg['To'] = ",".join(address_to)
msg['Subject']="Отчет по трафику юр.лиц".decode('utf-8')

msg.attach(part)
#text = "Во вложении прикреплен файл с юр.лицами, у которых за предыдущую неделю входящий трафик менее 50Мбайт"
part1 = MIMEText(text, 'plain')
msg.attach(part1)

# Send mail

server=smtplib.SMTP('mail.company.ru', 25)
#smtp.login(address, 'password')
server.sendmail(address_from, address_to, msg.as_string())
server.quit()
