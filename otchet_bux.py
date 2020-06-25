# -*- coding: utf-8 -*-
import smtplib
import xlsxwriter
import MySQLdb
import datetime
import os
import sys
import csv

from smtplib import SMTP_SSL
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders
from email.mime.text import MIMEText
from datetime import timedelta

reload(sys)
sys.setdefaultencoding('utf-8')

db = MySQLdb.connect('localhost','username','password','name_database')
db.set_character_set('utf8')
cur = db.cursor()
i=0
date=[
        ['Январь', '2020-01-01 10:00:00', '2020-02-01 03:00:00'],
        ['Февраль', '2020-02-01 10:00:00', '2020-03-01 03:00:00'],
        ['Март', '2020-03-01 10:00:00', '2020-04-01 03:00:00']
#        ['Апрель', '2019-04-01 10:00:00', '2019-05-01 03:00:00'],
#        ['Май', '2019-05-01 10:00:00', '2019-06-01 03:00:00'],
#        ['Июнь', '2019-06-01 10:00:00', '2019-07-01 03:00:00'],
#	['Июль', '2019-07-01 10:00:00', '2019-08-01 03:00:00'],
#	['Август', '2019-08-01 10:00:00', '2019-09-01 03:00:00'],
#	['Сентябрь', '2019-09-01 10:00:00', '2019-10-01 03:00:00'],
#	['Октябрь', '2019-10-01 10:00:00', '2019-11-01 03:00:00'],
#	['Ноябрь', '2019-11-01 10:00:00', '2019-12-01 03:00:00'],
#	['Декабрь', '2019-12-01 10:00:00', '2020-01-01 03:00:00'] '''
        ]

workbook = xlsxwriter.Workbook('otchet_bux_fiz1.xlsx')
worksheet = workbook.add_worksheet()
while i<(len(date)):
	cur.execute("""SELECT COUNT(users.full_name), SUM(invoice_entry.sum_cost) as summa
        	       FROM invoice_entry, invoices, users
        	       WHERE is_juridical=0
		       AND users.id=invoices.uid
        	       AND invoices.id=invoice_entry.invoice_id
        	       AND invoices.invoice_date>UNIX_TIMESTAMP('""" + date[i][1] + """')
        	       AND invoices.invoice_date<UNIX_TIMESTAMP('""" + date[i][2] + """')
		       AND invoice_entry.name REGEXP '.*Интернет.*'""")
	inet=cur.fetchone()
	print(inet[0])
	cur.execute("""SELECT COUNT(users.full_name),  SUM(invoice_entry.sum_cost) as summa
                       FROM invoice_entry, invoices, users
                       WHERE is_juridical=0
                       AND users.id=invoices.uid
                       AND invoices.id=invoice_entry.invoice_id
                       AND invoices.invoice_date>UNIX_TIMESTAMP('""" + date[i][1] + """')
                       AND invoices.invoice_date<UNIX_TIMESTAMP('""" + date[i][2] + """')
                       AND invoice_entry.name REGEXP '.*Телефония.*'
		       AND invoice_entry.name NOT REGEXP '.*внутризоновой.*'""")
	tel=cur.fetchone()
	print(tel[0])
	worksheet.write(i, 0, date[i][0])
	worksheet.write(i, 1, inet[0])
	worksheet.write(i, 2, inet[1])
	worksheet.write(i, 5, tel[0])
	worksheet.write(i, 6, tel[1])
	i+=1

workbook.close()
