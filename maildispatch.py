#! /usr/bin/python
# -*- coding: UTF-8 -*-
# Have to install modules xlrd, xlwt
#
# Go to this link and select Turn On
# to allow your gmail use unregistered app
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
import time

# Settings
excel = 'raw.xlsx' # path to excel file
image = u'attach.jpg' # PAth to image
sender = 'email@gmail.com' # From who email
replyto = 'email@gmail.com' # where to reply
login = 'email@gmail.com' # Login for gmail account
password = 'password' # Password for gmail account
subject = 'Subject of mail' # Subject of message

# Path to excel file (current in folder with script)
rb = xlrd.open_workbook(excel)
# Current active sheet
sheet = rb.sheet_by_index(0)
# Get all raw data from excel file
raw = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
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
for x in data:
	img = dict(title=u'Picture', path=image, cid=str(uuid.uuid4()))
	msg = MIMEMultipart('related')
	msg['Subject'] = subject # Subject of message
	msg['From'] = sender # From who
	msg['To'] = x[2] # To
	msg['Reply-to'] = replyto # Where to reply, should be the same as From
	msg_alternative = MIMEMultipart('alternative')
	msg.attach(msg_alternative)
	msg_text = MIMEText(u'[image: {title}]'.format(**img), 'plain', 'utf-8')
	msg_alternative.attach(msg_text)
	# Message body
	# tag <img has settings for image width in message (current 60%)
	# I'm not sure that it works correct, but you can play with width and find best value.
	msg_html = MIMEText(u'<div>Hello, '+x[0]+',</div><br>'
                     	 '<div>Hope you are doing well</div><br>'
                     	 '<div>example '+x[1]+' example</div><br>'
                     	 '<div>example</div>'
						 '<div dir="ltr">'
                     	 '<img src="cid:{cid}" alt="{alt}" width="60%"><br></div>'
                    	.format(alt=cgi.escape(img['title'], quote=True), **img),
                    	'html', 'utf-8')
	msg_alternative.attach(msg_html)
	with open(img['path'], 'rb') as file:
	    msg_image = MIMEImage(file.read(), name=os.path.basename(img['path']))
	    msg.attach(msg_image)
	msg_image.add_header('Content-ID', '<{}>'.format(img['cid']))

	# Create mail server
	server = smtplib.SMTP("smtp.gmail.com:587")
	server.ehlo()
	server.starttls()
	server.login(login, password) # Login and password from gmail account (Support only gmail for now)
	server.sendmail(msg['From'], x[2] , msg.as_string())
	server.quit()
	# Below just tech info for printing on the screen
	print 'Message for \033[33m%s\033[0m from \033[33m%s\033[0m to \033[33m%s\033[0m was \033[32mSENT\033[0m' % (x[0], x[1], x[2]) # print status on screen
	time.sleep(5) # wait 5 sec



