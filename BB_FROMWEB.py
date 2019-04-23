from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlwt
import xlrd
import time
from xlutils.copy import copy
import datetime
import datetime as dt
from tkinter import filedialog
from tkinter import *
from bs4 import BeautifulSoup
from urllib.request import urlopen
import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook


chrome_options = webdriver.ChromeOptions()
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options.add_experimental_option("prefs", prefs)

print ('▀▄▀▄▀▄ STOP ▄▀▄▀▄▀')



#-----EXCEL RESULT OPEN AND READ-----

book = xlrd.open_workbook(path_RESULT.filename)
wb=copy(book)
sheet_write = wb.get_sheet(0)
sheet_read = book.sheet_by_index(0)

wbxn = load_workbook(path_RESULT.filename)
wsn = wbxn.active

#-------FIND COLUMN UPDATE------
up=0
i=1
while up==0:
	#V_up=sheet_read.cell(0,i).value
	V_up=wsn.cell(row=1, column=i).value
	if V_up=='UPDATE_CALENDAR':
		up=1
	else:
		i=i+1
print(i)

#-------EMAIL VALUE-----------

#s = smtplib.SMTP('smtp.gmail.com', 587)
#s.starttls()
#s.login(sender, sender_password)

#-----RECUP INFO XPATH FROM EXCEL------
#book_XPATH = xlrd.open_workbook('/home/pi/Documents/PV/XPATH_.xls')
#sheet_XPATH = book_XPATH.sheet_by_index(0)
#X_comment=sheet_XPATH.cell(17,2).value
#XP_update=sheet_XPATH.cell(32,2).value
#XP_mois_1=sheet_XPATH.cell(30,2).value
#XP_mois_2=sheet_XPATH.cell(31,2).value
#XP_mois_3=sheet_XPATH.cell(33,3).value
#XP_next=sheet_XPATH.cell(34,2).value
#XP_CMP=sheet_XPATH.cell(35,2).value

#-------DATE DU JOUR-------
date = int(datetime.datetime.now().day)
month = int(datetime.datetime.now().month)
Hr=dt.datetime.now().hour

#------RECUP INFO CALANDAR------

def email(DIR2,NAMEFile,now):
	sender = 'stopairbnb.vincent@gmail.com'
	sender_password = '@STOP94AIRBNB'
	receivers = 'vincent.aulnay@gmail.com'

	s = smtplib.SMTP('smtp.gmail.com', 587)
	s.starttls()
	s.login(sender, sender_password)
	msg = MIMEMultipart()
	msg['From'] = sender
	msg['To'] = receivers
	#msg['Subject'] = "Subject of the Mail- image -2"
	body = "Body_of_the_mail"
	msg.attach(MIMEText(body, 'plain'))
	msg['Subject'] = "STOP AIRBNB - extraction du - "+str(now)
	# path along with extension of file to be attachmented 
	filename = DIR2+NAMEFile+str(now)+".xlsx"
	attachmentment = open(filename, "rb")
	 
	# instance of MIMEBase and named as p
	attachment = MIMEBase('application', 'octet-stream')
	# To change the payload into encoded form
	attachment.set_payload((attachmentment).read())
	# encode into base64
	encoders.encode_base64(attachment)
	attachment.add_header('Content-Disposition', "attachmentment; filename= %s" % filename)
	# attachment the instance  to instance 'msg'
	msg.attach(attachment)
	text = msg.as_string()
	s.sendmail(sender, receivers, text)
	print('*** email sent ***') 
	time.sleep(10)
	del filename
	del attachmentment
	del attachment
	del text
	del msg

def MnumDay (Mmois):
	global MNumday
	if Mmois=='janvier':
		MNumday=31
	elif Mmois=='février':
		MNumday=28
	elif Mmois=='mars':
		MNumday=31	
	elif Mmois=='avril':
		MNumday=30
	elif Mmois=='mai':
		MNumday=31
	elif Mmois=='juin':
		MNumday=30
	elif Mmois=='juillet':
		MNumday=31
	elif Mmois=='août':
		MNumday=30
	elif Mmois=='septembre':
		MNumday=31
	elif Mmois=='octobre':
		MNumday=30
	elif Mmois=='novembre':
		MNumday=31
	elif Mmois=='décembre':
		MNumday=30
		
		
		
		
def A_Colonne_mois(name_mois,c):
#1- récupération book Result qui évolue au court du script
#2- compter le nombre de colonne
#3- déterminer si colonne == name_mois de airbnb
#4- si condition alors c_write=c pour définir la colonne où écrire
	global c_write
	global new_month
	book_mois = xlrd.open_workbook(path_RESULT.filename)
	sheet_mois = book_mois.sheet_by_index(0)
	nc=sheet_mois.ncols
	#wbx = load_workbook(path_RESULT.filename)
	#ws = wbx.active
	
	new_month=0
	
	find_month=0
	while find_month==0:
		this_month=ws.cell(row=1, column=c+1).value
		#print(this_month)
		#Cell_MA=sheet_mois.cell(0,c-7).value
		#Cell_MA=ws.cell(row=c+1, column=c+2).value
		if this_month==name_mois:
			c_write=c+1
			break
		elif this_month==None:
			ws.cell(row=1, column=c+1).value = name_mois
			ws.cell(row=1, column=c+2).value = 'calendar y a 3 mois'
			ws.cell(row=1, column=c+3).value = 'jours disponible y a 3 mois'
			ws.cell(row=1, column=c+4).value = 'total réservé'
			ws.cell(row=1, column=c+5).value = 'NB_Comment'
			ws.cell(row=1, column=c+6).value = 'DIF_Comment'
			ws.cell(row=1, column=c+7).value = 'DIF_Nuitée'
			ws.cell(row=1, column=c+8).value = 'SOM_Nuitée'
			c_write=c+1
			find_month=1
			new_month=1
			print ('plus une colonne')
			wbx.save(path_RESULT.filename)
			break
		else:
			c=c+1


def A_Statu_day2(date,c_write,page,j,g,ResAirbnb,new_mo,MNday):	
	int_timeday=int(date)
	month=soup.findAll('div', attrs={"class":u"_1lds9wb"})[g]
	i=0
	li=[]
	if new_mo==1:
		ResAirbnb=''
	while i<=31:
		try:
			the_tr= month.findAll('td', attrs={"class": "_z39f86g"})[i]
			div=the_tr.find('div', attrs={"class": "_1fhupg9r"}).text
			intdiv=int(div)
			if intdiv>=int_timeday:
				li.append(intdiv)
			i=i+1
		except:
			break
	print (li)
	try:
		if len(li)>0:
			ca=ws.cell(row=j, column=c_write).value
			print(ca)
			#-------DATE DU JOUR-------
			date = int(datetime.datetime.now().day)
			month = int(datetime.datetime.now().month)
			toto=str(date)+'-'+str(month)
			if ca!=None:
				li_ca=ca.split(";")
			else:
				li_ca=[]

			lie=[]
			if li_ca!=[]:
				lenL=len(li_ca)
				h=0
				LB=[]
				while h!=lenL:
					LA=li_ca[h]
					LA=LA.split(':')
					del LA[0]
					LA=LA[0].split(',')
					lenLA=len(LA)
					g=0
					while g!=lenLA:
						intV=int(LA[g])
						LB.append(intV)
						g=g+1
					h=h+1
			
				lie=[elem for elem in li if elem not in LB ]
				if len(lie)!=0:
					#identification si nuitée est bloquée par préavis automatique
					preavis=''
					if len(lie)==1:
						dif=lie[0]-date
						preavis=''
						if dif==0 or dif==1 or dif==2 or dif==6:
							preavis='/P'
						elif dif<0:
							difP=MNday-date+lie[0]
							if difP==0 or difP==1 or difP==2 or difP==6:
								preavis='/P'
							
					t=ResAirbnb+preavis+toto+':'+str(lie)
					t=t.replace("[","")
					t=t.replace("]","")
					r=str(ca)+';    '+t
					lenli=len(lie)+len(LB)
					ws.cell(row=j, column=c_write+3).value=lenli
			else:
				t=ResAirbnb+toto+':'+str(li)
				t=t.replace("[","")
				t=t.replace("]","")
				r=t
				#print(r)
				lenli=len(li)
				ws.cell(row=j, column=c_write+3).value=lenli
			if r!='set()':
				print (r)
				ws.cell(row=j, column=c_write).value=r
	except:
		#print('rater 1')
		pass
	#COMMENTAIRE
	try:
		Bcomment=soup.find('button', attrs={"class": "_ff6jfq"})
		Scomment=Bcomment.find('span', attrs={"class": "_so3dpm2"}).text
		ws.cell(row=j, column=c_write+4).value=Scomment
	except:
		print('NO COMMENT')
		pass
		#wbx.save(path_RESULT.filename)

def A_Statu_day3(date,c_write,j):	
	int_timeday=int(date)
	
	month4=soup.find('div', attrs={"class":u"_kuxo8ai"})
	i=0
	li=[]
	while i<=31:
		try:
			the_tr= month4.findAll('td', attrs={"class": "_12fun97"})[i]
			div=the_tr.find('div', attrs={"class": "_1tpncgrb"}).text
			intdiv=int(div)
			if intdiv>=int_timeday:
				li.append(intdiv)
			i=i+1
		except:
			break
	
	liste=[]
	liste=li
	liste.sort()
	liste=set(liste)
	lenli=len(liste)
	#sheet_write.write(j,c_write+2,lenli)
	ws.cell(row=j, column=c_write+2).value=lenli
	print (liste)
	strli=str(liste)
	str_repl_1=strli.replace("'","")
	str_repl_2=str_repl_1.replace("{","") #
	str_repl_3=str_repl_2.replace("}","") #
	if str_repl_1!='set()':
		#sheet_write.write(j,c_write+1,str_repl_3)
		ws.cell(row=j, column=c_write+1).value=str_repl_3
	#wb.save(path_RESULT.filename)
	
def A_Statu_day4(c_write,j,ResAirbnb,new_mo):	
	month5=soup.find('div', attrs={"class":u"_kuxo8ai"})
	i=0
	li=[]
	if new_mo==1:
		ResAirbnb=''
	while i<=31:
		try:
			the_tr= month5.findAll('td', attrs={"class": "_z39f86g"})[i]
			div=the_tr.find('div', attrs={"class": "_1fhupg9r"}).text
			intdiv=int(div)
			li.append(intdiv)
			i=i+1
		except:
			break
	try:
		if len(li)>0:
			#book_date = xlrd.open_workbook(path_RESULT.filename)
			#sheet_date = book_date.sheet_by_index(0)
			#ca=sheet_date.cell(j,c_write).value
			ca=ws.cell(row=j, column=c_write).value
			#-------DATE DU JOUR-------
			date = int(datetime.datetime.now().day)
			month = int(datetime.datetime.now().month)
			toto=str(date)+'-'+str(month)
			if ca!=None:
				li_ca=ca.split(";")
			else:
				li_ca=[]

			lie=[]
			if li_ca!=[]:
				lenL=len(li_ca)
				h=0
				LB=[]
				while h!=lenL:
					LA=li_ca[h]
					LA=LA.split(':')
					del LA[0]
					LA=LA[0].split(',')
					lenLA=len(LA)
					g=0
					while g!=lenLA:
						intV=int(LA[g])
						LB.append(intV)
						g=g+1
					h=h+1
			
				lie=[elem for elem in li if elem not in LB ]
				if len(lie)!=0:
					t=ResAirbnb+toto+':'+str(lie)
					t=t.replace("[","")
					t=t.replace("]","")
					r=str(ca)+';    '+t
					lenli=len(lie)+len(LB)
					ws.cell(row=j, column=c_write+3).value=lenli
			else:
				t=ResAirbnb+toto+':'+str(li)
				t=t.replace("[","")
				t=t.replace("]","")
				r=t
				#print(r)
				lenli=len(li)
				ws.cell(row=j, column=c_write+3).value=lenli
			if r!='set()':
				print (r)
				sheet_write.write(j,c_write,r)
				ws.cell(row=j, column=c_write).value=r
	except:
		pass
	#wb.save(path_RESULT.filename)
	
def Colonne_mois(name_mois,c,year):
#1- récupération book Result qui évolue au court du script
#2- compter le nombre de colonne
#3- déterminer si colonne == name_mois de airbnb
#4- si condition alors c_write=c pour définir la colonne où écrire
	global c_write
	global new_month
	book_mois = xlrd.open_workbook(path_RESULT.filename)
	sheet_mois = book_mois.sheet_by_index(0)
	nc=sheet_mois.ncols
	#wbx = load_workbook(path_RESULT.filename)
	#ws = wbx.active
	
	new_month=0
	
	find_month=0
	month_year=name_mois+'_'+year
	while find_month==0:
		this_month=ws.cell(row=1, column=c+1).value
		#print(this_month)
		#Cell_MA=sheet_mois.cell(0,c-7).value
		#Cell_MA=ws.cell(row=c+1, column=c+2).value
		if this_month==month_year:
			c_write=c+1
			break
		elif this_month==None:
			ws.cell(row=1, column=c+1).value = month_year
			ws.cell(row=1, column=c+2).value = 'calendar y a 3 mois'
			ws.cell(row=1, column=c+3).value = 'jours disponible y a 3 mois'
			ws.cell(row=1, column=c+4).value = 'total réservé'
			ws.cell(row=1, column=c+5).value = 'NB_Comment'
			ws.cell(row=1, column=c+6).value = 'DIF_Comment'
			ws.cell(row=1, column=c+7).value = 'DIF_Nuitée'
			ws.cell(row=1, column=c+8).value = 'SOM_Nuitée'
			c_write=c+1
			find_month=1
			new_month=1
			print ('plus une colonne')
			wbx.save(path_RESULT.filename)
			break
		else:
			c=c+1


def Statu_day2(date,c_write,page,j,g,ResAirbnb,new_mo):	
	int_timeday=int(date)
	month=soup.findAll('tbody', attrs={"class":"day-wrap"})[g]
	i=0
	li=[]
	if new_mo==1:
		ResAirbnb=''
	while i<=31:
		try:
			the_tr= month.findAll('div', {"class": re.compile("pm-unavailable")})[i]
			div=the_tr.find('div', attrs={"class": "day-template__day"}).text
			intdiv=int(div)
			if intdiv>=int_timeday:
				li.append(intdiv)
			i=i+1
		except:
			break
	
	try:
		if len(li)>0:
			ca=ws.cell(row=j, column=c_write).value
			#-------DATE DU JOUR-------
			date = int(datetime.datetime.now().day)
			month = int(datetime.datetime.now().month)
			toto=str(date)+'-'+str(month)
			if ca!=None:
				li_ca=ca.split(";")
			else:
				li_ca=[]

			lie=[]
			if li_ca!=[]:
				lenL=len(li_ca)
				h=0
				LB=[]
				while h!=lenL:
					LA=li_ca[h]
					LA=LA.split(':')
					del LA[0]
					LA=LA[0].split(',')
					lenLA=len(LA)
					g=0
					while g!=lenLA:
						intV=int(LA[g])
						LB.append(intV)
						g=g+1
					h=h+1
			
				lie=[elem for elem in li if elem not in LB ]
				if len(lie)!=0:
					t=ResAirbnb+toto+':'+str(lie)
					t=t.replace("[","")
					t=t.replace("]","")
					r=str(ca)+';    '+t
					lenli=len(lie)+len(LB)
					ws.cell(row=j, column=c_write+3).value=lenli
			else:
				t=ResAirbnb+toto+':'+str(li)
				t=t.replace("[","")
				t=t.replace("]","")
				r=t
				#print(r)
				lenli=len(li)
				ws.cell(row=j, column=c_write+3).value=lenli
			if r!='set()':
				print (r)
				ws.cell(row=j, column=c_write).value=r
	except:
		#print('rater 1')
		pass
	#COMMENTAIRE
	try:
		Bcomment=soup.find('h2', attrs={"class": "review-summary__header-overview-headline"})
		Scomment=Bcomment.find('span').text
		Pcomment=Scomment.split(' ')
		ws.cell(row=j, column=c_write+4).value=Pcomment[0]
	except:
		print('NO COMMENT')
		pass
		#wbx.save(path_RESULT.filename)

def Statu_day3(date,c_write,j):	
	int_timeday=int(date)
	
	#month4=soup.find('div', attrs={"class":u"_kuxo8ai"})
	month4=soup.findAll('tbody', attrs={"class":"day-wrap"})[2]
	i=0
	li=[]
	while i<=31:
		try:
			#the_tr= month4.findAll('td', attrs={"class": "_12fun97"})[i]
			#div=the_tr.find('div', attrs={"class": "_1tpncgrb"}).text
			the_tr= month4.findAll('div', {"class": re.compile("day-template--available-stay")})[i]
			div=the_tr.find('div', attrs={"class": "day-template__day"}).text
			intdiv=int(div)
			if intdiv>=int_timeday:
				li.append(intdiv)
			i=i+1
		except:
			break
	
	liste=[]
	liste=li
	liste.sort()
	liste=set(liste)
	lenli=len(liste)
	#sheet_write.write(j,c_write+2,lenli)
	ws.cell(row=j, column=c_write+2).value=lenli
	print (liste)
	strli=str(liste)
	str_repl_1=strli.replace("'","")
	str_repl_2=str_repl_1.replace("{","") #
	str_repl_3=str_repl_2.replace("}","") #
	if str_repl_1!='set()':
		#sheet_write.write(j,c_write+1,str_repl_3)
		ws.cell(row=j, column=c_write+1).value=str_repl_3
	#wb.save(path_RESULT.filename)
	
def Statu_day4(c_write,j,ResAirbnb,new_mo):	
	#month5=soup.find('div', attrs={"class":u"_kuxo8ai"})
	month5=soup.findAll('tbody', attrs={"class":"day-wrap"})[2]
	i=0
	li=[]
	if new_mo==1:
		ResAirbnb=''
	while i<=31:
		try:
			#the_tr= month5.findAll('td', attrs={"class": "_z39f86g"})[i]
			#div=the_tr.find('div', attrs={"class": "_1rcgiovb"}).text
			the_tr= month5.findAll('div', {"class": re.compile("pm-unavailable")})[i]
			div=the_tr.find('div', attrs={"class": "day-template__day"}).text
			intdiv=int(div)
			li.append(intdiv)
			i=i+1
		except:
			break
	try:
		if len(li)>0:
			#book_date = xlrd.open_workbook(path_RESULT.filename)
			#sheet_date = book_date.sheet_by_index(0)
			#ca=sheet_date.cell(j,c_write).value
			ca=ws.cell(row=j, column=c_write).value
			#-------DATE DU JOUR-------
			date = int(datetime.datetime.now().day)
			month = int(datetime.datetime.now().month)
			toto=str(date)+'-'+str(month)
			if ca!=None:
				li_ca=ca.split(";")
			else:
				li_ca=[]

			lie=[]
			if li_ca!=[]:
				lenL=len(li_ca)
				h=0
				LB=[]
				while h!=lenL:
					LA=li_ca[h]
					LA=LA.split(':')
					del LA[0]
					LA=LA[0].split(',')
					lenLA=len(LA)
					g=0
					while g!=lenLA:
						intV=int(LA[g])
						LB.append(intV)
						g=g+1
					h=h+1
			
				lie=[elem for elem in li if elem not in LB ]
				if len(lie)!=0:
					t=ResAirbnb+toto+':'+str(lie)
					t=t.replace("[","")
					t=t.replace("]","")
					r=str(ca)+';    '+t
					lenli=len(lie)+len(LB)
					#sheet_write.write(j,c_write+3,lenli)
					ws.cell(row=j, column=c_write+3).value=lenli
			else:
				t=ResAirbnb+toto+':'+str(li)
				t=t.replace("[","")
				t=t.replace("]","")
				r=t
				#print(r)
				lenli=len(li)
				#sheet_write.write(j,c_write+3,lenli)
				ws.cell(row=j, column=c_write+3).value=lenli
			if r!='set()':
				print (r)
				sheet_write.write(j,c_write,r)
				ws.cell(row=j, column=c_write).value=r
	except:
		pass
	#wb.save(path_RESULT.filename)

	
#-----OPEN GOOGLE CHROME and AIRBNB PAGE---------

#rootdriver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
rootdriver = webdriver.Chrome(chrome_options=chrome_options)
#rootdriver.set_page_load_timeout(2)
rootdriver.set_window_size(2000, 1000)
wait = WebDriverWait(rootdriver, 3)

#nrow=(sheet_read.nrows)+1

nrow=wsn.max_row
print('NROW'+str(nrow))
j=2
z=0
end=0
C_mois=0
EE=0
Tr=0
j=2
z=0
end=0
C_mois=0
date = int(datetime.datetime.now().day)
#Hr=dt.datetime.now().hour
wbx = load_workbook(path_RESULT.filename)
ws = wbx.active
while end==0:
	try:
		while j<=nrow:
			print('-------------')
			print (j-1)
			h=ws.cell(row=j, column=2).value
			print(h)
			if 'airbnb' in h:
				rootdriver.get(h)
				time.sleep(1)
				html = rootdriver.page_source
				soup = BeautifulSoup(html, 'html.parser')
				ResAirbnb=''
				try:
					update=soup.find('div', attrs={"class":u"_q401y8m"})
					V_up=update.find('span').text
					ws.cell(row=j, column=i).value=V_up
					wbx.save(path_RESULT.filename)
					if V_up!="Mis à jour aujourd'hui":
						ResAirbnb='/A'
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 1--------
					if C_mois==0:
						name_mois1 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_1lds9wb'][1]//div[@class='_gucugi']/strong"))).text
						print(name_mois1)
						Mname1=name_mois1.split(' ')
						MN1=Mname1[0]
						run_MN=MnumDay(MN1)
						print (MNumday)
						MNday1=MNumday
						run_c=A_Colonne_mois(name_mois1,i)
						m1_write=c_write
						m1_newmonth=new_month
					print('   ---')
					print('le mois N est '+name_mois1)
					run_day=A_Statu_day2(date,m1_write,1,j,0,ResAirbnb,m1_newmonth,500)
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 2--------
					if C_mois==0:
						name_mois2 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='_1lds9wb'][2]//div[@class='_gucugi']/strong"))).text
						print(name_mois2)
						Mname2=name_mois2.split(' ')
						MN2=Mname2[0]
						run_MN=MnumDay(MN2)
						print (MNumday)
						MNday2=MNumday
						run_c=A_Colonne_mois(name_mois2,i)
						m2_write=c_write
						m2_newmonth=new_month
					print('   ---')
					print('le mois N+1 est '+name_mois2)
					run_day=A_Statu_day2(1,m2_write,2,j,1,ResAirbnb,m2_newmonth,MNday1)
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 3--------
					print('   ---')
					if C_mois==0:
						month31=soup.findAll('div', attrs={"class":u"_gucugi"})[3]
						name_mois3=month31.find('strong').text
						run_c=A_Colonne_mois(name_mois3,i)
						m3_write=c_write
						m3_newmonth=new_month
					#if z==0: #check si nuit dispo en vide dans excel
						#bookW = xlrd.open_workbook(path_RESULT.filename)
						#sheet_readW = bookW.sheet_by_index(0)
						#c=(sheet_readW.ncols)-6
						#c=ws.cell(row=j, column=c_write+2).value
					#d=sheet_readW.cell(j,c).value
					d=ws.cell(row=j, column=m3_write+2).value
					#if d==None:
					#	print('le mois N+2 est '+name_mois3)
					#	run_day=A_Statu_day3(1,m3_write,j)
					#	run_resday=A_Statu_day4(m3_write,j,ResAirbnb,m3_newmonth)
					#else:
					print('le mois N+2 est '+name_mois3)
					run_resday=A_Statu_day4(m3_write,j,ResAirbnb,m3_newmonth)
					#print('Jours disponible déjà capturés')
				except:
					print('PAS DE MOIS 3')
					pass
				wbx.save(path_RESULT.filename)
				C_mois=1
				j=j+1
			elif 'abritel' in h:
				rootdriver.get(h)
				time.sleep(1)
				html = rootdriver.page_source
				soup = BeautifulSoup(html, 'html.parser')
				ResAirbnb=''
				try:
					oldV_up=ws.cell(row=j, column=i).value
					update=soup.find('dl', attrs={"data-key":u"availabilityUpdated"})
					V_up=update.find('dt').text
					ws.cell(row=j, column=i).value=V_up
					wbx.save(path_RESULT.filename)
					if V_up==oldV_up:
						ResAirbnb='/A'
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 1--------
					#if C_mois==0:
					#	name_mois1 = wait.until(EC.presence_of_element_located((By.XPATH, XP_mois_1))).text
					#	year_mois1 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'calendar') and contains(@class ,'learfix')]/div[1]/h4/span[2]"))).text
					#	print(name_mois1+' - '+year_mois1)
					#	run_c=Colonne_mois(name_mois1,i,year_mois1)
					#	m1_write=c_write
					#	m1_newmonth=new_month
					print('   ---')
					print('le mois N est '+name_mois1)
					run_day=Statu_day2(date,m1_write,1,j,0,ResAirbnb,m1_newmonth)
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 2--------
					#if C_mois==0:
					#	name_mois2 = wait.until(EC.presence_of_element_located((By.XPATH, XP_mois_2))).text
					#	year_mois2 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'calendar') and contains(@class ,'learfix')]/div[2]/h4/span[2]"))).text
					#	run_c=Colonne_mois(name_mois2,i,year_mois2)
					#	m2_write=c_write
					#	m2_newmonth=new_month
					print('   ---')
					print('le mois N+1 est '+name_mois2)
					run_day=Statu_day2(1,m2_write,2,j,1,ResAirbnb,m2_newmonth)
				except:
					pass
				try:
				#-----RECUPERATION CALANDAR MOIS 3--------
					print('   ---')
					#if C_mois==0:
					#	name_mois3 = wait.until(EC.presence_of_element_located((By.XPATH, XP_mois_3))).text
					#	year_mois3 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'calendar') and contains(@class ,'learfix')]/div[3]/h4/span[2]"))).text
					#	print(name_mois3+' - '+year_mois1)
					#	run_c=Colonne_mois(name_mois3,i,year_mois3)
					#	m3_write=c_write
					#	m3_newmonth=new_month
					#if z==0: #check si nuit dispo en vide dans excel
						#bookW = xlrd.open_workbook(path_RESULT.filename)
						#sheet_readW = bookW.sheet_by_index(0)
						#c=(sheet_readW.ncols)-6
						#c=ws.cell(row=j, column=c_write+2).value
					#d=sheet_readW.cell(j,c).value
					#d=ws.cell(row=j, column=m3_write+2).value
					print('le mois N+2 est '+name_mois3)
					#run_day=Statu_day3(1,m3_write,j)
					run_resday=Statu_day4(m3_write,j,ResAirbnb,m3_newmonth)
				except:
					print('PAS DE MOIS 3')
					pass
				wbx.save(path_RESULT.filename)
				C_mois=1
				j=j+1
		
		end=1
		now = str(datetime.datetime.now())[:19]
		now = now.replace(":","_")
		wbx.save(DIR2+NAMEFile+str(now)+".xlsx")
		run=email(DIR2,NAMEFile,now)
		Tr=date
		print ('_______    ___    ___     ___')
		print ('|      |   |  |   |  \    |  |')
		print ('|  |__     |  |   |   \   |  |')
		print ('|     |    |  |   |    \  |  |')
		print ('|  |       |  |   |  |\ \ |  |')
		print ('|  |       |  |   |  | \ \|  |')
		print ('|__|       |__|   |__|  \____|')
		rootdriver.quit()
	except:
		# EXCEPT si Chrome se ferme tout seul, ici il va le réouvrir et relancer la boucle d'extraction
		#rootdriver = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver',chrome_options=chrome_options)
		rootdriver = webdriver.Chrome(chrome_options=chrome_options)
		rootdriver.set_window_size(1000, 1500)
		wait = WebDriverWait(rootdriver, 3)


print('FIN')
exist()






	






