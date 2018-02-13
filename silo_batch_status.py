import os
from configparser import SafeConfigParser
import datetime
import logging
import pyodbc
import time
from bs4.builder import HTML


#Logging Configurations
logging.basicConfig(filename='status.log',level=logging.DEBUG, format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p')

#Directory Configurations
current_working_dir = os.getcwd()

#Static Variables Declarations
server = ''
database = ''
table = ''
silos = []
statement = ''
to_address = []
cc_address = []

#Processs Variables Declarations
check_log = True
query = ''
silo_previous_count = []
silo_current_count = []	
processed_since = ''
display_note = False
log_message = ''
html_message_body = ''
zero_batch_silos = []
zero_batch = False
message_body = ''

def mail_status():
	global to_address
	global cc_address
	global html_message_body
	import win32com.client as win32
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail_to = ";".join(to_address)
	mail.To = mail_to
	mail_cc=";".join(cc_address)
	mail.CC = mail_cc
	mail.Subject = 'Silo Monitoring Status - ' + str(datetime.date.today().strftime('%m/%d/%Y'))
	mail.HTMLBody = html_message_body
	try:
		mail.send()
	except Exception:
		print('Exception')	
	
def build_mail_body():
	global silo_current_count
	global html_message_body
	global zero_batch
	global zero_batch_silos
	html_table = "<table border = "'1'" cellpadding="'10'">"
	html_table+="<tr bgcolor="'#2196F3'">"
	html_table+="<td align="'center'">" + "Silo" +"</td>"
	html_table+="<td align="'center'">" + "Batches" +"</td>"
	html_table+="<td align="'center'">" + "Processed" +"</td>"
	html_table+="</tr>"
	for a in silo_current_count:
		html_table+="<tr>"
		for key in a.keys():	
			if key == 'Silo':
				html_table+="<td align="'center'">Silo "+str(a[key])+"</td>"
			else:
				html_table+="<td align="'center'">"+str(a[key])+"</td>"
		html_table+="</tr>"
	html_table+="</table>"
	html_message_body +='<html><body>Hi Team,<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find the Silo Status below,<br><br>'
	if zero_batch:
		html_message_body+='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Silos with zero batches : ' + ", ".join(zero_batch_silos) + " (Load Batches)"
	if check_log:
		html_message_body += '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Processed Since : ' + str(processed_since) + '<br><br>'
	html_message_body += html_table
	if display_note:
		html_message_body += '<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Note: New Batches added / Log File Modified / Silos Changed'
	html_message_body += '<br><br>Thanks,<br><br>Offshore Imaging Support</body></html>'		

def build_log():
	global log_message
	for single_silo in silo_current_count:
		log_message += "silo=" + str(single_silo['Silo']) + " " + "batches=" + str(single_silo['Count']) + " " + "processed=" + str(single_silo['Processed']) + ";"

def log():
	global log_message
	build_log()
	logging.info(log_message)

def calculate_status():
	global silo_current_count
	global silo_previous_count
	global zero_batch_silos
	global zero_batch
	for current_silo in silo_current_count:
		for previous_silo in silo_previous_count:
			if current_silo['Silo'] == previous_silo['Silo']:
				priv = int(previous_silo['Count'])
				cur = int(current_silo['Count'])
				if cur == 0:
					zero_batch = True
					zero_batch_silos.append(current_silo['Silo'])
				silos_processed = priv - cur
				if silos_processed < 0:
					silos_processed = 0
					display_note = True
				current_silo['Processed'] = silos_processed
				del(silos_processed)
				del(priv)
				del(cur)

def read_log_file():
	global processed_since
	log_file = open('status.log','r')
	for line in log_file:
		continue
	processed_since = line[0:23]
	previous_status = line[23:].split(";")
	for single_status_data in previous_status[0:len(previous_status)-1]:
		single_status = single_status_data.split(" ")
		silo = single_status[0].split("=")[1]
		for previous_silo in silo_previous_count:
			if previous_silo['Silo'] == silo:
				previous_silo['Count'] = single_status[1].split("=")[1]
	
def execute_query(silo):
	global query
	query_to_execute = query.replace('<silo>',str(silo))
	results = cur.execute(query_to_execute)
	for result in results:
		count = result[0]
		break;
	return count;

def get_status():
	for silo in silos:
		status = {}
		status['Silo'] = silo
		status['Count'] = execute_query(silo)
		status['Processed'] = '0'
		silo_current_count.append(status)
		del(status)

def build_query():
	global statement
	global query
	query = statement 
	query = query.replace('<server>',server)
	query = query.replace('<database>',database)
	query = query.replace('<table>',table)

def process():
	build_query()
	get_status()
	calculate_status()
	log()
	build_mail_body()
	mail_status()
	if zero_batch:
		mail_zero_batch()
	

def mock_log_file():
	for silo in silos:
		status = {}
		status['Silo'] = silo
		status['Count'] = '0'
		silo_previous_count.append(status)
		del(status)

def isLogEmpty():
	if os.stat("status.log").st_size == 0:
		return True
	else:
		return False

def fetch_declarations():
	global server
	global database
	global table
	global silos
	global statement
	global to_address
	global cc_address
	
	parser = SafeConfigParser()
	parser.read(current_working_dir+"\\config.ini")
	 
	#Fetching Query Related Details
	server = parser.get('query', 'server')
	database = parser.get('query', 'database')
	table = parser.get('query', 'table')
	silos = parser.get('query', 'silos').split(',')
	statement = parser.get('query', 'statement')
	 
	#fetching Mail Addresses
	to_address = parser.get('mail_address', 'to_address').split(';')
	cc_address = parser.get('mail_address', 'cc_address').split(';')
	
print("-----------------")
print("Silo Batch Status")
print("-----------------")
fetch_declarations();

#Establishing Connection With Database
trim_server = server[1:len(server)-1]
trim_database = database[1:len(database)-1]
con = pyodbc.connect(driver="{SQL Server}",server=trim_server,database=trim_database,Trusted_Connection='yes')
cur = con.cursor()

mock_log_file()

if not os.path.isfile('status.log') or isLogEmpty():
	print("Log File Does Not Exists (or) Empty")
	check_log = False
	
else:
	print("Log File exists\n")
	check_log = True
	read_log_file()

process()