import os
from configparser import SafeConfigParser
import datetime
import time
import fnmatch
import xlsxwriter
from bs4.builder import HTML

#Declarations
cur_wrk_dir = os.getcwd()
folder_paths = []
corelogic_files = []
table_html = ''
count = 1
	
def archive():
	excel_name=now.replace("/","-")+".xlsx"
	workbook = xlsxwriter.Workbook(cur_wrk_dir+"\\Archive\\"+excel_name)
	worksheet = workbook.add_worksheet()
	row = 0
	col = 0
	worksheet.write(row,col,"S No")
	col+=1
	worksheet.write(row,col,"Folder")
	col+=1
	worksheet.write(row,col,"Filename")
	col=0
	row+=1
	for dict in corelogic_files:
		for key in dict.keys():
			worksheet.write(row,col,dict[key])
			col+=1
		row+=1
		col=0

#Step 15 - Getting the mailIds from the file
def get_mail_list(file_name):
	mail_list=[]
	file=open(file_name,"r")
	for line in file:
		if ".com" in line:	
			mail_list.append(line)
	return mail_list

#Step 14 - Fetch the mailing list
def set_mail_list(type):
	if type == "to":
		mail_path=os.getcwd()+"\\Utility\\to_address.txt"
		mail_list=get_mail_list(mail_path)
		return mail_list
	else:
		mail_path=os.getcwd()+"\\Utility\\cc_address.txt"
		mail_list=get_mail_list(mail_path)
		return mail_list

def sendTableMail(table_message_body):
	import win32com.client as win32
	to_mail_list=[]
	cc_mail_list=[]
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	to_mail_list=set_mail_list("to")
	mail_to=";".join(to_mail_list)
	mail.To = mail_to
	to_mail_list=set_mail_list("cc")
	mail_cc=";".join(to_mail_list)
	mail.CC = mail_cc
	sub="SFTP Process Notification - Recordation Nationstar File Transfer "+ now
	mail_subject=sub
	mail.Subject = mail_subject
	mail.HTMLBody = table_message_body
	try:
		mail.send()
	except Exception:
		pass

#Create status table
def create_table():
	global corelogic_files
	global table_html
	table_html+="<table border = "'1'" cellpadding="'10'">"
	for a in corelogic_files:
		table_html+="<tr bgcolor="'#2196F3'">"
		for key in a.keys():
			table_html+="<td align="'center'">"+key.upper()+"</td>"
		table_html+="</tr>"
		break
	for a in corelogic_files:
		table_html+="<tr>"
		for key in a.keys():
			if key == "S No" or str(a[key]) == "-":
				table_html+="<td align="'center'">"+str(a[key])+"</td>"
			else:
				table_html+="<td>"+str(a[key])+"</td>"
		table_html+="</tr>"
	table_html+="</table>"

#Get File names
def fetch_files(path,tag):
	global count
	temp = {}
	temp['S No'] = count
	count+=1
	temp['Folder'] = tag
	temp['Filename'] =	'-'
	for f in os.listdir(path):
		if fnmatch.fnmatch(f, '*'+now1+'*.zip'):
				temp['Filename'] =	str(f)
	corelogic_files.append(temp)
	del(temp)
	
#Get file count
def get_file_count(path):
	temp = 0
	for f in os.listdir(path):
		if fnmatch.fnmatch(f, '*'+now+'*.zip'):
				temp+=1
	return temp

#Fetch folder paths
def fetch_folder_paths():
	parser = SafeConfigParser()
	parser.read(cur_wrk_dir+"\\Utility\\config.ini")
	folder_paths.append(parser.get('paths', 'unexecuted_modification_documents'))
	folder_paths.append(parser.get('paths', 'modification_recording_rejection'))
	folder_paths.append(parser.get('paths', 'mod_agreement'))
	folder_paths.append(parser.get('paths', 'modification_fha_trial_agreement'))

print("----------------------------- ")
print("Corelogic Transfer Monitoring ")
print("----------------------------- ")
choice = str(input("\nProcess for:\n\n1)Today\n\n2)Yesterday\n\n3)Enter a date\n\nYour Input....."))
if choice == "1":
	now = str(datetime.date.today().strftime('%m/%d/%Y'))
	now1 = now.replace("/","")
if choice == "2":
	now = str((datetime.date.today()- datetime.timedelta(days=1)).strftime('%m/%d/%Y'))
	now1 = now.replace("/","")
elif choice == "3":
	scount=0
	now = input("\nEnter date in MM/DD/YYYY format:")
	if now[1] == '/' or now[2] != '/' or now[5] != '/' or int(now[0:2]) <= 0 or int(now[0:2]) > 12 or int(now[3:5]) <= 0 or int(now[3:5]) > 31 or int(now[6:]) <= 2000:
		print("\nInvalid input date format...quitting the process")
		b=input("\nPress enter to exit")
		if b != '':
			exit()
		
	now1 = now.replace("/","")

else:
	print("\nInvalid input...quitting the process")
	b=input("\nPress enter to exit")
	if b != '':
		exit()
	

print("\nProcessing.....")

#Fetch folder paths - Main

fetch_folder_paths()

#Fetch zip files present in each path
fetch_files(folder_paths[0],"Unexecuted Modification Documents")
fetch_files(folder_paths[1],"Modification Recording Rejection(County Rejection)")
fetch_files(folder_paths[2],"Mod Agreement(Recorded)")
fetch_files(folder_paths[3],"Modification FHA Trial Agreement")
create_table()
table_message_body='<html><body>Hello Corelogic Team,<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;File Created Date :' + now + '<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;We have received the following files from Corelogic:<br><br><div style="margin-left:30px">'+table_html+'</div><br><br>Thanks,<br>Imaging Support</body></html>'
sendTableMail(table_message_body)
archive()
a=input("\nProcess completed!Press enter to quit the process")
if a != '':
	exit()