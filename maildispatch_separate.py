#! /usr/bin/python
# -*- coding: UTF-8 -*-
# Have to install modules xlrd, xlwt
#
# Go to this link and select Turn On
# https://www.google.com/settings/security/lesssecureapps

import xlrd, xlwt
import os
import cgi
import uuid
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image     import MIMEImage
from email.header         import Header
from smtplib import SMTP
import smtplib
import sys

rb = xlrd.open_workbook('raw.xlsx') # Path to excel file (current in folder with script)
sheet = rb.sheet_by_index(0) # Current active sheet
raw = [sheet.row_values(rownum) for rownum in range(sheet.nrows)] # Get all raw data from excel file
data = []

# Parse data
for row in raw:
	# if we have 'position' word in row then we pass it
	# to avoid first row with column names.
	# (Can be changed for other column name )
	if 'Position' in row:
		continue
	x = []
	name = row[2].split(' ')[0] # get name from column and cut only first name (just cut off al after fisrt space)
	x.append(name)
	x.append(row[4]) # get agency from 5 column.
	x.append(row[6]) # get email from 7 column.
	data.append(x)

# Create message
image = 'attach.jpg' # PAth to image
sender = 'email@gmail.com' # From who email
login = 'email@gmail.com' # Login for gmail account
password = 'password'
for x in data:
	msg = MIMEMultipart('related')
	msg['Subject'] = 'Subject of email'
	msg['From'] = sender # From who
	msg['Reply-to'] = sender # Where to reply, should be the same lika from
	# Message body (can be modified)
	msg_html = MIMEText(u'<div>Hello, '+x[0]+',</div><br>'
                     	 '<div>Hope you are doing well</div><br>'
                     	 '<div>Example of message</div><br>'
                     	 '<div>Cheers example</div>',
                    	'html', 'utf-8')
	msg.attach(msg_html)

	part = MIMEApplication(open(image,"rb").read())
	part.add_header('Content-Disposition', 'attachment', filename=image)
	msg.attach(part)

	# Create mail server
	server = smtplib.SMTP("smtp.gmail.com:587")
	server.ehlo()
	server.starttls()
	server.login(login, password) # Login and password from gmail account (Support only gmail for now)
	server.sendmail(msg['From'], x[2] , msg.as_string()) # Sending email
	server.quit()
	print 'Message for %s from %s to %s was SENT' % (x[0], x[1], x[2]) # print status on screen


