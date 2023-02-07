# import shutil
# from time import sleep
# from random import choice
# import ctypes
# import PILasda
# import ssl
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
from openpyxl.cell import cell
from openpyxl.chart import LineChart, Reference
# import xml.etree.ElementTree as ET
from werkzeug.utils import secure_filename
from flask import flash
from flask import Flask, render_template, request, send_from_directory, after_this_request, redirect, url_for
from openpyxl.descriptors import ( 
	String,
	Sequence,
	Integer,
	)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.styles import Alignment, alignment
# from string import ascii_uppercase
import openpyxl
# import datetime
# from datetime import datetime adsasda
import os
app=Flask(__name__)
app.secret_key = "GT ROMANIA Delivery Center"
var_list=[]
var_rute=[]
# app.app_context().push()
	
# @app.route("/")
# def home():
#     # text=request.form.get('client')
#     # print(text)
#     return render_template("index.html")


@app.route('/')
def FS():
	return render_template('FS.html')

@app.route('/Instructions', methods=['GET'])
def downloadPMG():
	# filepath = "D:\Projects\8. Python web apps\Test web flask\Instructions"
	return send_from_directory("/home/fsbot/storage","Instructions - FS.docx", as_attachment=True)
@app.route('/', methods=['POST', 'GET'])
def FS_process():
	company = request.form['company']
	address = request.form['address']
	vatTaxCode = request.form['code']
	regNr = request.form['registration']
	typeOfCompany = request.form['type']
	mainActivity = request.form['activity']
	year= request.form['year']
	dropdownlimba = request.form.get('limba')
	dropdownfroma = request.form.get('forma')

	if str(dropdownlimba)=="Romana(RO)":
		option=1
	else:
		option=0

	if str(dropdownfroma)=="Romana(RO)":
		option2=1
	else:
		option2=0


	folderpath="/home/fsbot/storage"
	# folderpath="C:\\Users\\denis.david\\Desktop\\Mirusv4\\auditAppsMirus\\Output\\JE"

	if request.method == 'POST':

		workingsblue2= Font(bold=True, italic=True, name='Tahoma', size=8,color='FFFFFF')
		lbluefill = PatternFill(start_color='7030A0',
							end_color='7030A0',
							fill_type='solid')
		grifill=PatternFill(start_color='c4d79b',end_color='c4d79b',fill_type='solid')
		yellow=PatternFill(start_color='ffff00',end_color='ffff00',fill_type='solid')
		blueFill = PatternFill(start_color='00AEAC',
							end_color='00AEAC',
							fill_type='solid')
		doubleborder = Border(bottom=Side(style='double'))
		solidborder = Border(bottom=Side(style='thick'))
		solidborderstanga = Border(left=Side(style='thin'))
		rightborder = Border(right=Side(style='thin'))
		rightdouble = Border (right=Side(style='thin'), bottom=Side(style='double'))
		rightmedium = Border (right=Side(style='thin'), bottom=Side(style='medium'))
		solidborderdreapta = Border(right=Side(style='thin'))
		solidbordersus = Border(top=Side(style='thin'))
		fontitalic = Font(name='Tahoma', size=8, bold=True, italic=True)
		font1 = Font(name='Tahoma', size=8)
		font2 = Font(name='Tahoma', size=8, bold=True)
		fontRed = Font(name='Tahoma', size=8, bold=True, color= 'FF0000')
		fontRedDiff=Font(name="Tahoma", color='FF0000', size=11, )
		fontGT = Font (name='GT Logo', size=8)
		workingsblue = Font(color='2F75B5', bold=True, name='Tahoma', size=8 )
		headers= Font(bold=True, italic=True, name='Tahoma', size=8,color='FFFFFF') 
		headersblue = PatternFill(start_color='7030A0',
						end_color='7030A0',
						fill_type='solid')
		headerspurple= PatternFill(start_color='65CDCC',
							end_color='65CDCC',
							fill_type='solid')
		total=PatternFill(start_color='DDD9C4',
						end_color='DDD9C4',
						fill_type='solid')
		greenbolditalic= Font(bold=True, italic=True,  color='C0504D', name='Tahoma', size=8)
		greenbolditalic= Font(bold=True, italic=True,  color='00af50')
		fontalb = Font(italic=True, color="bfbfbf", size=8, name='Tahoma')
		
		triald=request.files["TB"]
			

		# PBC_CY=mapping.create_sheet("Trial Balance2")
		# PBC_CY.sheet_view.showGridLines = False
		if(option2==0):
			if(option==0):
				mapping=openpyxl.load_workbook('/home/fsbot/exceltemp/SF Entitati mici_EN.xlsx')
				# mapping=openpyxl.load_workbook('C:\\Users\\denis.david\\Training materials\\SF Entitati mici_EN.xlsx')			

			else:
				mapping=openpyxl.load_workbook('/home/fsbot/exceltemp/SF Entitati mici_RO.xlsx')
				# mapping=openpyxl.load_workbook('C:\\Users\\denis.david\\Training materials\\SF Entitati mici_RO.xlsx')
			ws=mapping.active		
			TBCY = openpyxl.load_workbook(triald,data_only=True)
			TBCY1 = TBCY.active
			# PBC_CY=mapping.create_sheet("TB_PBC")
			test=mapping["Trial Balance"]
			test2=mapping["Check if manual ADJE"]


			for row in TBCY1.iter_rows():
					for cell in row:
						if cell.value=="Account":
							tbCyAcount=cell.column
							tbrow=cell.row

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="Description":
						tbCyDescription=cell.column

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="OB":
						tbCyOB=cell.column

			for row in TBCY1.iter_rows():

				for cell in row:
					if cell.value=="DM":
						tbCyDM=cell.column
					
			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="CM":
						tbCyCM=cell.column

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="CB":
						tbCyCB=cell.column



			try:
				luntb=len(TBCY1[tbCyAcount])
			except:
				flash("Please insert the correct header for Account in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
				# sys.exit()
			try:
				Account=[b.value for b in TBCY1[tbCyAcount][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for Account in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
				# sys.exit()

			try:
				Description=[b.value for b in TBCY1[tbCyDescription][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for Description in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Description'")
				# sys.exit()
			try:
				OB=[b.value for b in TBCY1[tbCyOB][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for OB")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Debit'")
				# sys.exit()
			try:
				DM=[b.value for b in TBCY1[tbCyDM][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for DM")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Credit'")
				# sys.exit()
			try:
				CM=[b.value for b in TBCY1[tbCyCM][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for CM")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Rulaj Curent Debit'")
				# sys.exit()
			try:
				CB=[b.value for b in TBCY1[tbCyCB][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for CB")
				return render_template("index.html")

			for i in range(1, len(Account)+1):
				test.cell(row=i+14, column=6).value=Account[i-1]

			for i in range (1, len(Description)+1):
				test.cell(row=i+14, column=7).value= Description[i-1]

			for i in range (1, len(OB)+1):
				test.cell(row=i+14, column=8).value=OB[i-1]

			for i in range (1, len(DM)+1):
				test.cell (row=i+14, column =9).value=DM[i-1]

			for i in range (1,len(CM)+1):
				test.cell (row=i+14, column=10).value=CM[i-1]

			for i in range (1,len(CB)+1):
				test.cell (row=i+14, column=11).value=CB[i-1]


			for i in range(1, len(Account)+1):
				test.cell(row=i+14,column=2).value='=_xlfn.NUMBERVALUE(LEFT(F{0},1))'.format(i+14)	
			for i in range(1, len(Account)+1):
				test.cell(row=i+14,column=1).value='=IF(B'+str(14+i)+'<6,"BS",IF(B'+str(14+i)+'=6,"Exp","Rev"))'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=3).value='=Left(F'+str(14+i)+',2)'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=4).value='=Left(F'+str(14+i)+',3)'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=5).value='=IF(F'+str(14+i)+'="121",Left(F'+str(14+i)+',3)&"0",Left(F'+str(14+i)+',4))'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=12).value='=K'+str(14+i)+'-H'+str(14+i)+''
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=13).value='=IFERROR(L'+str(14+i)+'/H'+str(14+i)+'," ")'
			for i in range(1,len(Account)+1):
				# test.cell(row=i+14,column=14).value='''=_xlfn.IF(A'''+str(14+i)+'''="BS"'''+''',IFERROR(VLOOKUP(TRIM($E'''+str(14+i)+'),'+"'BS Mapping std'"+'!$A:$D,4,0),VLOOKUP(TRIM($D'+str(14+i)+'),'+"'BS Mapping std'"+'!$A:$D,4,0)),IFNA(VLOOKUP(TRIM($E'+str(14+i)+'),'+"'PL mapping Std'"+'!$A:$D,4,0),VLOOKUP(TRIM($D'+str(14+i)+'),'+"'PL mapping Std'"+'!$A:$D,4,0)))'
				test.cell(row=i+14,column=14).value='''=IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$C,3,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$C,3,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$C,3,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$C,3,0)))'''.format(i+14)
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=15).value="=_xlfn.IFERROR(VLOOKUP(E"+str(14+i)+",'F30 mapping'!A:C,3,0),VLOOKUP(D"+str(14+i)+",'F30 mapping'!A:C,3,0))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=16).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:C,3,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:C,3,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=17).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:D,4,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:D,4,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=18).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:E,5,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:E,5,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=19).value="=_xlfn.IF(B"+str(14+i)+"<6,IFERROR(VLOOKUP(E"+str(14+i)+",'BS Mapping std'!A:F,6,0),VLOOKUP(D"+str(14+i)+",'BS Mapping std'!A:F,6,0)),IFERROR(VLOOKUP(E"+str(14+i)+",'PL mapping Std'!A:D,4,0),VLOOKUP(D"+str(14+i)+",'PL mapping Std'!A:D,4,0)))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=20).value="=_xlfn.IF(B"+str(14+i)+"<6,IFERROR(VLOOKUP(E"+str(14+i)+",'BS Mapping std'!A:G,7,0),VLOOKUP(D"+str(14+i)+",'BS Mapping std'!A:G,7,0)),IFERROR(VLOOKUP(E"+str(14+i)+",'PL mapping Std'!A:E,5,0),VLOOKUP(D"+str(14+i)+",'PL mapping Std'!A:E,5,0)))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=22).value='''=IF(IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$H,8,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$H,8,0)))=0,"",IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$H,8,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$H,8,0))))'''.format(i+14)
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=23).value="=_xlfn.IFERROR(VLOOKUP(E"+str(14+i)+",'F30 mapping'!A:D,4,0),VLOOKUP(D"+str(14+i)+",'F30 mapping'!A:D,4,0))"

			# for i in range(len(Account)+1,800):
			# 	test.cell(row=i+14,column=14).value=""
			# 	test.cell(row=i+14,column=15).value=""
			# 	test.cell(row=i+14,column=16).value=""
			# 	test.cell(row=i+14,column=17).value=""
			# 	test.cell(row=i+14,column=18).value=""
			# 	test.cell(row=i+14,column=19).value=""
			# 	test.cell(row=i+14,column=20).value=""
			# 	test.cell(row=i+14,column=21).value=""
			# 	test.cell(row=i+14,column=22).value=""
			# 	test.cell(row=i+14,column=23).value=""



			test.cell(row=1,column=2).value=company
			test.cell(row=2,column=2).value=address
			test.cell(row=3,column=2).value=vatTaxCode
			test.cell(row=4,column=2).value=regNr
			test.cell(row=5,column=2).value=typeOfCompany
			test.cell(row=6,column=2).value=mainActivity
			test.cell(row=7,column=2).value=year

			listaMapare=['161','1614','1615','1617','1618','1621','1622','1624','1625','1627','1623','1626','1661','1663','2671','2672','2673','2674','2675','2676','2677','2678','2679','308','348E','348','368','378','388','4428','445','4451','4452','4458','446','4482','4481','4511','4518','4531','4538','456','456','471','472','473','4751','4752','4753','4754','4758','481','482','581']
			listaBalanta=list(set(Account))
			listasetmapare=list(set(listaMapare))
			print(listaBalanta)
			print(listaMapare)
			val=[]
			for i in range(0,len(listaBalanta)):
				for j in range(0,len(listasetmapare)):
					if(str(listaBalanta[i])[0:3]==str(listasetmapare[j]) or str(listaBalanta[i])[0:4]==str(listasetmapare[j])):
						val.append(listaBalanta[i])

			for j  in range(0,len(val)):
				test2.cell(row=14+j,column=6).value=val[j]
			for x in range(0,len(val)):
				test2.cell(row=14+x,column=5).value='=Left(F'+str(14+x)+',4)'
				test2.cell(row=14+x,column=4).value='=Left(F'+str(14+x)+',3)'
				test2.cell(row=14+x,column=3).value='=Left(F'+str(14+x)+',2)'
				test2.cell(row=14+x,column=2).value='=Left(F'+str(14+x)+',1)'
				test2.cell(row=14+x,column=1).value='BS'
				test2.cell(row=14+x,column=7).value="=VLOOKUP(F{0},'Trial Balance'!F:G,2,0)".format(x+14)
				test2.cell(row=14+x,column=8).value="=VLOOKUP(F{0},'Trial Balance'!F:H,3,0)".format(x+14)
				test2.cell(row=14+x,column=9).value="=VLOOKUP(F{0},'Trial Balance'!F:I,4,0)".format(x+14)
				test2.cell(row=14+x,column=10).value="=VLOOKUP(F{0},'Trial Balance'!F:J,5,0)".format(x+14)
				test2.cell(row=14+x,column=11).value="=VLOOKUP(F{0},'Trial Balance'!F:K,6,0)".format(x+14)
				test2.cell(row=x+14,column=12).value='=K'+str(14+x)+'-H'+str(14+x)+''
				test2.cell(row=x+14,column=13).value='=IFERROR(L'+str(14+x)+'/H'+str(14+x)+'," ")'
				test2.cell(row=x+14,column=14).value='''=IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$D,4,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$D,4,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$D,4,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$D,4,0)))'''.format(x+14)
				# test2.cell(row=14+x,column=14).value='Bifunctional - Please asses'
					# else:
					# 	print("nu sunt la fel")


			print(len(listaBalanta))
			print(len(listaMapare))

			f10=mapping["1. F10"]
			f20=mapping["2. F20"]
			f30=mapping["3. F30"]
			f40=mapping["4. F40"]
			# soce=mapping["5. SOCE"]
			# socf=mapping["6.SOCF"]
			n3nca=mapping["N3 - NCA"]
			n4inv=mapping["N4 - Inventories"]
			n7cash=mapping["N7 - Cash"]
			n5tr=mapping["N5 - TR"]
			n9tp=mapping["N9 - TP"]
			n10prov=mapping["N10 - Provisions"]
			n15pers=mapping["N15 - Personnel"]
			n16opex=mapping["N16 - Other OPEX"]

			f10.print_area="A11:E132"
			f20.print_area="A11:E91"
			f30.print_area="A17:F292"
			f40.print_area="A10:H79"
			# soce.print_area="E8:L38"
			# socf.print_area="A7:C51"
			n3nca.print_area="A10:O73"
			n4inv.print_area="A10:G24"
			n5tr.print_area="A10:F47"
			n7cash.print_area="A10:C19"
			n9tp.print_area="A10:G41"
			n10prov.print_area="A10:G22"
			n15pers.print_area="A10:C27"
			n16opex.print_area="A10:C30"
			mapping.save(str(folderpath)+"/Financial Statements-"+str(company)+".xlsx")
			# return send_from_directory(folderpath,"Financial Statements-"+str(company)+".xlsx",as_attachment=True)

		else:
			if(option==0):
				mapping=openpyxl.load_workbook('/home/fsbot/exceltemp/Template FS ENG.xlsx')
				# mapping=openpyxl.load_workbook('C:\\Users\\denis.david\\Training materials\\Template FS ENG.xlsx')			

			else:
				mapping=openpyxl.load_workbook('/home/fsbot/exceltemp/Template FS RO.xlsx')
				# mapping=openpyxl.load_workbook('C:\\Users\\denis.david\\Training materials\\Template FS RO.xlsx')
			ws=mapping.active		
			TBCY = openpyxl.load_workbook(triald)
			TBCY1 = TBCY.active
			# PBC_CY=mapping.create_sheet("TB_PBC")
			test=mapping["Trial Balance"]
			test2=mapping["Check if manual ADJE"]


			for row in TBCY1.iter_rows():
					for cell in row:
						if cell.value=="Account":
							tbCyAcount=cell.column
							tbrow=cell.row

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="Description":
						tbCyDescription=cell.column

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="OB":
						tbCyOB=cell.column

			for row in TBCY1.iter_rows():

				for cell in row:
					if cell.value=="DM":
						tbCyDM=cell.column
					
			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="CM":
						tbCyCM=cell.column

			for row in TBCY1.iter_rows():
				for cell in row:
					if cell.value=="CB":
						tbCyCB=cell.column



			try:
				luntb=len(TBCY1[tbCyAcount])
			except:
				flash("Please insert the correct header for Account in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
				# sys.exit()
			try:
				Account=[b.value for b in TBCY1[tbCyAcount][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for Account in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
				# sys.exit()

			try:
				Description=[b.value for b in TBCY1[tbCyDescription][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for Description in Trial Balance file")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Description'")
				# sys.exit()
			try:
				OB=[b.value for b in TBCY1[tbCyOB][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for OB")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Debit'")
				# sys.exit()
			try:
				DM=[b.value for b in TBCY1[tbCyDM][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for DM")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Credit'")
				# sys.exit()
			try:
				CM=[b.value for b in TBCY1[tbCyCM][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for CM")
				return render_template("index.html")
				# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Rulaj Curent Debit'")
				# sys.exit()
			try:
				CB=[b.value for b in TBCY1[tbCyCB][tbrow:luntb+1]]
			except:
				flash("Please insert the correct header for CB")
				return render_template("index.html")

			for i in range(1, len(Account)+1):
				test.cell(row=i+14, column=6).value=str(Account[i-1])

			for i in range (1, len(Description)+1):
				test.cell(row=i+14, column=7).value= Description[i-1]

			for i in range (1, len(OB)+1):
				test.cell(row=i+14, column=8).value=OB[i-1]

			for i in range (1, len(DM)+1):
				test.cell (row=i+14, column =9).value=DM[i-1]

			for i in range (1,len(CM)+1):
				test.cell (row=i+14, column=10).value=CM[i-1]

			for i in range (1,len(CB)+1):
				test.cell (row=i+14, column=11).value=CB[i-1]


			for i in range(1, len(Account)+1):
				test.cell(row=i+14,column=2).value='=_xlfn.NUMBERVALUE(LEFT(F{0},1))'.format(i+14)	
			for i in range(1, len(Account)+1):
				test.cell(row=i+14,column=1).value='=IF(B'+str(14+i)+'<6,"BS",IF(B'+str(14+i)+'=6,"Exp","Rev"))'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=3).value='=Left(F'+str(14+i)+',2)'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=4).value='=Left(F'+str(14+i)+',3)'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=5).value='=IF(F'+str(14+i)+'="121",Left(F'+str(14+i)+',3)&"0",Left(F'+str(14+i)+',4))'
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=12).value='=K'+str(14+i)+'-H'+str(14+i)+''
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=13).value='=IFERROR(L'+str(14+i)+'/H'+str(14+i)+'," ")'
			for i in range(1,len(Account)+1):
				# test.cell(row=i+14,column=14).value='''=_xlfn.IF(A'''+str(14+i)+'''="BS"'''+''',IFERROR(VLOOKUP(TRIM($E'''+str(14+i)+'),'+"'BS Mapping std'"+'!$A:$D,4,0),VLOOKUP(TRIM($D'+str(14+i)+'),'+"'BS Mapping std'"+'!$A:$D,4,0)),IFNA(VLOOKUP(TRIM($E'+str(14+i)+'),'+"'PL mapping Std'"+'!$A:$D,4,0),VLOOKUP(TRIM($D'+str(14+i)+'),'+"'PL mapping Std'"+'!$A:$D,4,0)))'
				test.cell(row=i+14,column=14).value='''=IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$D,4,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$D,4,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$D,4,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$D,4,0)))'''.format(i+14)
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=15).value="=_xlfn.IFERROR(VLOOKUP(E"+str(14+i)+",'F30 mapping'!A:C,3,0),VLOOKUP(D"+str(14+i)+",'F30 mapping'!A:C,3,0))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=16).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:C,3,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:C,3,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=17).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:D,4,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:D,4,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=18).value="=_xlfn.IFERROR(IFERROR(VLOOKUP(E"+str(14+i)+",'F40 mapping'!A:E,5,0),VLOOKUP(D"+str(14+i)+",'F40 mapping'!A:E,5,0)),0)"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=19).value="=_xlfn.IF(B"+str(14+i)+"<6,IFERROR(VLOOKUP(E"+str(14+i)+",'BS Mapping std'!A:E,5,0),VLOOKUP(D"+str(14+i)+",'BS Mapping std'!A:E,5,0)),IFERROR(VLOOKUP(E"+str(14+i)+",'PL mapping Std'!A:F,6,0),VLOOKUP(D"+str(14+i)+",'PL mapping Std'!A:F,6,0)))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=20).value="=_xlfn.IF(B"+str(14+i)+"<6,IFERROR(VLOOKUP(E"+str(14+i)+",'BS Mapping std'!A:F,6,0),VLOOKUP(D"+str(14+i)+",'BS Mapping std'!A:F,6,0)),IFERROR(VLOOKUP(E"+str(14+i)+",'PL mapping Std'!A:G,7,0),VLOOKUP(D"+str(14+i)+",'PL mapping Std'!A:G,7,0)))"
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=22).value='''=IF(IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$H,8,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$H,8,0)))=0,"",IF(A{0}="BS",IFERROR(VLOOKUP(TRIM($E{0}),'BS Mapping std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'BS Mapping std'!$A:$H,8,0)),IFERROR(VLOOKUP(TRIM($E{0}),'PL mapping Std'!$A:$H,8,0),VLOOKUP(TRIM($D{0}),'PL mapping Std'!$A:$H,8,0))))'''.format(i+14)
			for i in range(1,len(Account)+1):
				test.cell(row=i+14,column=23).value="=_xlfn.IFERROR(VLOOKUP(E"+str(14+i)+",'F30 mapping'!A:D,4,0),VLOOKUP(D"+str(14+i)+",'F30 mapping'!A:D,4,0))"

			# for i in range(len(Account)+1,800):
			# 	test.cell(row=i+14,column=14).value=""
			# 	test.cell(row=i+14,column=15).value=""
			# 	test.cell(row=i+14,column=16).value=""
			# 	test.cell(row=i+14,column=17).value=""
			# 	test.cell(row=i+14,column=18).value=""
			# 	test.cell(row=i+14,column=19).value=""
			# 	test.cell(row=i+14,column=20).value=""
			# 	test.cell(row=i+14,column=21).value=""
			# 	test.cell(row=i+14,column=22).value=""
			# 	test.cell(row=i+14,column=23).value=""



			test.cell(row=1,column=2).value=company
			test.cell(row=2,column=2).value=address
			test.cell(row=3,column=2).value=vatTaxCode
			test.cell(row=4,column=2).value=regNr
			test.cell(row=5,column=2).value=typeOfCompany
			test.cell(row=6,column=2).value=mainActivity
			test.cell(row=7,column=2).value=year

			listaMapare=['161','1614','1615','1617','1618','1621','1622','1624','1625','1627','1623','1626','1661','1663','2671','2672','2673','2674','2675','2676','2677','2678','2679','308','348E','348','368','378','388','4428','445','4451','4452','4458','446','4482','4481','4511','4518','4531','4538','456','456','471','472','473','4751','4752','4753','4754','4758','481','482','581']
			listaBalanta=list(set(Account))
			listasetmapare=list(set(listaMapare))
			print(listaBalanta)
			print(listaMapare)
			val=[]
			for i in range(0,len(listaBalanta)):
				for j in range(0,len(listasetmapare)):
					if(str(listaBalanta[i])[0:3]==str(listasetmapare[j]) or str(listaBalanta[i])[0:4]==str(listasetmapare[j])):
						val.append(listaBalanta[i])

			for j  in range(0,len(val)):
				test2.cell(row=14+j,column=6).value=val[j]
			for x in range(0,len(val)):
				test2.cell(row=14+x,column=5).value='=Left(F'+str(14+x)+',4)'
				test2.cell(row=14+x,column=4).value='=Left(F'+str(14+x)+',3)'
				test2.cell(row=14+x,column=3).value='=Left(F'+str(14+x)+',2)'
				test2.cell(row=14+x,column=2).value='=Left(F'+str(14+x)+',1)'
				test2.cell(row=14+x,column=1).value='BS'
				test2.cell(row=14+x,column=7).value="=VLOOKUP(F{0},'Trial Balance'!F:G,2,0)".format(x+14)
				test2.cell(row=14+x,column=8).value="=VLOOKUP(F{0},'Trial Balance'!F:H,3,0)".format(x+14)
				test2.cell(row=14+x,column=9).value="=VLOOKUP(F{0},'Trial Balance'!F:I,4,0)".format(x+14)
				test2.cell(row=14+x,column=10).value="=VLOOKUP(F{0},'Trial Balance'!F:J,5,0)".format(x+14)
				test2.cell(row=14+x,column=11).value="=VLOOKUP(F{0},'Trial Balance'!F:K,6,0)".format(x+14)
				test2.cell(row=x+14,column=12).value='=K'+str(14+x)+'-H'+str(14+x)+''
				test2.cell(row=x+14,column=13).value='=IFERROR(L'+str(14+x)+'/H'+str(14+x)+'," ")'
				test2.cell(row=x+14,column=14).value="=VLOOKUP(F{0},'Trial Balance'!F:N,9,0)".format(x+14)
				test2.cell(row=14+x,column=15).value='Bifunctional - Please asses'
					# else:
					# 	print("nu sunt la fel")


			print(len(listaBalanta))
			print(len(listaMapare))

			f10=mapping["1. F10"]
			f20=mapping["2. F20"]
			f30=mapping["3. F30"]
			f40=mapping["4. F40"]
			soce=mapping["5. SOCE"]
			socf=mapping["6.SOCF"]
			n3nca=mapping["N3 - NCA"]
			n4inv=mapping["N4 - Inventories"]
			n5tr=mapping["N5 - TR"]
			n7cash=mapping["N7 - Cash"]
			n9tp=mapping["N9 - TP"]
			n10prov=mapping["N10 - Provisions"]
			n15pers=mapping["N15 - Personnel"]
			n16opex=mapping["N16 - Other OPEX"]

			f10.print_area="A11:E132"
			f20.print_area="A11:E91"
			f30.print_area="A17:F292"
			f40.print_area="A10:H79"
			soce.print_area="E8:L38"
			socf.print_area="A7:C51"
			n3nca.print_area="A10:O73"
			n4inv.print_area="A10:G24"
			n5tr.print_area="A10:F47"
			n7cash.print_area="A10:C19"
			n9tp.print_area="A10:G41"
			n10prov.print_area="A10:G22"
			n15pers.print_area="A10:C27"
			n16opex.print_area="A10:C30"
			mapping.save(str(folderpath)+"/Financial Statements-"+str(company)+".xlsx")
	return send_from_directory(folderpath,"Financial Statements-"+str(company)+".xlsx",as_attachment=True)



if __name__ == '__main__':
   	app.run()
