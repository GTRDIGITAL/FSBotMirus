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
import datetime
from datetime import datetime
import os
from string import ascii_uppercase
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
	# folderpath="D:\\FSFinal\\FSBotMirus"

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


@app.route('/TrialBalances/Instructions', methods=['GET'])
def downloadTB():
		filepath = "/home/fsbot/storage/Instructions - Trial Balance.docx"
 
		return send_from_directory(filepath,"Instructions - Trial Balance.docx", as_attachment=True)
@app.route('/TrialBalances/GT3SjGyxpbcxV35PeSUpKJQIOgY')
def TB():
	return render_template('TB.html')
@app.route('/TrialBalances/GT3SjGyxpbcxV35PeSUpKJQIOgY', methods=['POST', 'GET'])
def TB_process():

	namec = request.form['client']
	ant= datetime.strptime(
					 request.form['yearEnd'],
					 '%Y-%m-%d')
	threshol = request.form['threshold']
	preparedBy1 = request.form['preparedBy']
	isChecked1=request.form.get("Stdmapp")
	isChecked2=request.form.get("forml")
	isChecked3=request.form.get("forms")
	isChecked4=request.form.get("pyEx")
	# denis=datetime.datetime.now()


		# if isChecked1=="": #daca e bifat
	#     isChecked=1
	# else:
	#     isChecked=0
	
	# folderpath="/home/auditappnexia/output/tb"
	folderpath="/home/fsbot/storage"

	def make_archive(source, destination):
		base = os.path.basename(destination)
		name = base.split('.')[0]
		format = base.split('.')[1]
		archive_from = os.path.dirname(source)
		archive_to = os.path.basename(source.strip(os.sep))
		shutil.make_archive(name, format, archive_from, archive_to)
		shutil.move('%s.%s'%(name,format), destination)
	# yearEnd = str(request.form['yearEnd'])
	# processed_text = client.upper()
	# fisier=request.files.get('monthlyTB')
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
		font = Font(name='Tahoma', size=8, bold=True)
		font1 = Font(name='Tahoma', size=8)
		font2 = Font(name='Tahoma', size=10, bold=True)
		fontRed = Font(name='Tahoma', size=10, bold=True, color= 'FF0000')
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
		trialc=request.files["trialBalCYPBC"]
		trialp=request.files["trialBalPYPBC"]
		TBCY = openpyxl.load_workbook(trialc,data_only=True)
		TBCY1 = TBCY.active

		# if isChecked4=="":
		# 	try:
		
		"Open files"




		"Iterate from imported PBC's:"


		'Iterate from CY TB'

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
				if cell.value=="SID":
					tbCySID=cell.column

		for row in TBCY1.iter_rows():

			for cell in row:
				if cell.value=="SIC":
					tbCySIC=cell.column
				
		for row in TBCY1.iter_rows():
			for cell in row:
				if cell.value=="RCD":
					tbCyRCD=cell.column

		for row in TBCY1.iter_rows():
			for cell in row:
				if cell.value=="RCC":
					tbCyRCC=cell.column

		for row in TBCY1.iter_rows():
			for cell in row:
				if cell.value=="SFD":
					tbCySFD=cell.column

		for row in TBCY1.iter_rows():
			for cell in row:
				if cell.value=="SFC":
					tbCySFC=cell.column


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
			SID=[b.value for b in TBCY1[tbCySID][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Sold Initial Debit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Debit'")
			# sys.exit()
		try:
			SIC=[b.value for b in TBCY1[tbCySIC][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Sold Initial Credit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Initial Credit'")
			# sys.exit()
		try:
			RCD=[b.value for b in TBCY1[tbCyRCD][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Rulaj Curent Debit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Rulaj Curent Debit'")
			# sys.exit()
		try:
			RCC=[b.value for b in TBCY1[tbCyRCC][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Rulaj Curent Credit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Rulaj Curent credit'")
			# sys.exit()
		try:
			SFD=[b.value for b in TBCY1[tbCySFD][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Sold Final Debit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Final Debit'")
			# sys.exit()
		try: 
			SFC=[b.value for b in TBCY1[tbCySFC][tbrow:luntb+1]]
		except:
			flash("Please insert the correct header for Sold Final Credit in Trial Balance file")
			return render_template("index.html")
			# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Final Credit'")
			# sys.exit()
		"Create CY PBC"

		if isChecked4=="":
			TBPY = openpyxl.load_workbook(trialp,data_only=True)
			TBPY1 = TBPY.active
			try:
				for row in TBPY1.iter_rows():
						for cell in row:
							if cell.value=="Account":
								tbPyAcount=cell.column
								tbPYrow=cell.row

				for row in TBPY1.iter_rows():
					for cell in row:
						if cell.value=="Description":
							tbPyDescription=cell.column


				for row in TBPY1.iter_rows():
					for cell in row:
						if cell.value=="CB":
							tbPySFD=cell.column


				try:
					luntbp=len(TBPY1[tbPyAcount])
				except:
					flash("Please insert the correct header for Account in Trial Balance Prior Year file")
					return render_template("index.html")
					# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
					# sys.exit()
				try:
					Accountp=[b.value for b in TBPY1[tbPyAcount][tbPYrow:luntb+1]]
				except:
					flash("Please insert the correct header for Account in Trial Balance Prior Year file")
					return render_template("index.html")
					# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Account'")
					# sys.exit()

				try:
					Descriptionp=[b.value for b in TBPY1[tbPyDescription][tbPYrow:luntb+1]]
				except:
					flash("Please insert the correct header for Description in Trial Balance Prior Year file")
					return render_template("index.html")
					# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Description'")
					# sys.exit()

				try:
					CB=[b.value for b in TBPY1[tbPySFD][tbPYrow:luntb+1]]
				except:
					flash("Please insert the correct header for CB in Trial Balance Prior Year file")
					return render_template("index.html")
					# messagebox.showerror("Error", "File: Trial Balance. Please insert the correct header for 'Sold Final Debit'")
					# sys.exit()
			except:
				pass
		
		# if isChecked1=="":
		# 	mapp=request.files["Mapping"]
		output=openpyxl.Workbook()
		# else:
		# 	if isChecked2=="":
		# 		output=openpyxl.load_workbook("/home/auditappnexia/output/otherfiles/Mapping Forma Lunga.xlsx",data_only=True)
		# 	else:
		# 		output=openpyxl.load_workbook("/home/auditappnexia/output/otherfiles/Mapping Forma Scurta.xlsx",data_only=True)

		PBC_CY =output.create_sheet("PBC_CY")

		PBC_CY.cell(row=1, column=1).value="Class"
		PBC_CY.cell(row=1, column=2).value="Synt3"
		PBC_CY.cell(row=1, column=3).value="Synt4"
		PBC_CY.cell(row=1, column=4).value="Account"
		PBC_CY.cell(row=1, column=5).value="Description"
		PBC_CY.cell(row=1, column=6).value="SID"
		PBC_CY.cell(row=1, column=7).value="SIC"
		PBC_CY.cell(row=1, column=8).value="RCD"
		PBC_CY.cell(row=1, column=9).value="RCC"
		PBC_CY.cell(row=1, column=10).value="SFD"
		PBC_CY.cell(row=1, column=11).value="SFC"


		for i in range (1,10):
			PBC_CY.cell (row=1, column=i).border=doubleborder
			PBC_CY.cell (row=1, column=i).font=font2


		for i in range(1, len(Account)+1):
			PBC_CY.cell(row=i+1, column=4).value=Account[i-1]

		for i in range (1, len(Description)+1):
			PBC_CY.cell(row=i+1, column=5).value= Description[i-1]

		for i in range (1, len(SID)+1):
			PBC_CY.cell(row=i+1, column=6).value=SID[i-1]

		for i in range (1, len(SIC)+1):
			PBC_CY.cell (row=i+1, column =7).value=SIC[i-1]

		for i in range (1,len(RCD)+1):
			PBC_CY.cell (row=i+1, column=8).value=RCD[i-1]

		for i in range (1,len(RCC)+1):
			PBC_CY.cell (row=i+1, column=9).value=RCC[i-1]

		for i in range (1,len(SFD)+1):
			PBC_CY.cell (row=i+1, column=10).value=SFD[i-1]

		for i in range(1,len(SFC)+1):
			PBC_CY.cell (row=i+1, column=11).value=SFC[i-1]

		for i in range (1,12):
			PBC_CY.cell(row=1, column=i).font=font2
			PBC_CY.cell(row=1, column=i).border=doubleborder
			PBC_CY.cell(row=1, column=i).fill=blueFill

		for i in range (1, len(SFD)+1):
			for j in range (6, 12):
				PBC_CY.cell(row=i+1, column=j).number_format='#,##0_);(#,##0)'


		for i in range(1, len(Account)+1):
			PBC_CY.cell(row=i+1,column=1).value='=Left(D{0},1)'.format(i+1)
		for i in range(1,len(Account)+1):
			PBC_CY.cell(row=i+1,column=2).value='=Left(D{0},3)'.format(i+1)
		for i in range(1,len(Account)+1):
			PBC_CY.cell(row=i+1,column=3).value='=Left(D{0},4)'.format(i+1)


		PBC_PY =output.create_sheet("PBC_PY")

		if isChecked4=="":
			try:
				PBC_PY.cell(row=1, column=1).value="Class"
				PBC_PY.cell(row=1, column=2).value="Synt3"
				PBC_PY.cell(row=1, column=3).value="Synt4"
				PBC_PY.cell(row=1, column=4).value="Account"
				PBC_PY.cell(row=1, column=5).value="Description"
				PBC_PY.cell(row=1, column=6).value="CB"


				for i in range (1,8):
					PBC_PY.cell (row=1, column=i).border=doubleborder
					PBC_PY.cell (row=1, column=i).font=font2


				for i in range(1, len(Accountp)+1):
						PBC_PY.cell(row=i+1, column=4).value=Accountp[i-1]

				for i in range (1, len(Descriptionp)+1):
					PBC_PY.cell(row=i+1, column=5).value= Descriptionp[i-1]


				for i in range (1,len(CB)+1):
					PBC_PY.cell (row=i+1, column=6).value=CB[i-1]

				for i in range (1,8):
					PBC_PY.cell(row=1, column=i).font=font2
					PBC_PY.cell(row=1, column=i).border=doubleborder
					PBC_PY.cell(row=1, column=i).fill=blueFill

				for i in range (1, len(CB)+1):
					for j in range (6, 8):
						PBC_PY.cell(row=i+1, column=j).number_format='#,##0_);(#,##0)'


				for i in range(1, len(Accountp)+1):
					PBC_PY.cell(row=i+1,column=1).value='=Left(D{0},1)'.format(i+1)
				for i in range(1,len(Accountp)+1):
					PBC_PY.cell(row=i+1,column=2).value='=Left(D{0},3)'.format(i+1)
				for i in range(1,len(Accountp)+1):
					PBC_PY.cell(row=i+1,column=3).value='=Left(D{0},4)'.format(i+1)
			except:
				pass

		"Define F10 Worksheet"

		F10TB=output.create_sheet("F_10_Trial_Balance")
		F10TB.sheet_view.showGridLines = False
		F10TB.cell(row=1, column=1).value="Client:"
		F10TB.cell(row=1, column=1).font=font
		F10TB.cell(row=1, column=2).value=namec
		F10TB.cell(row=1, column=2).font=font


		F10TB.cell(row=2, column=1).value="Period end:"
		F10TB.cell(row=2, column=1).font=font
		F10TB.cell(row=2, column=2).value=ant
		F10TB.cell(row=2, column=2).font=font
		F10TB.cell(row=2, column=2).number_format="mm/dd/yyyy"


		F10TB.cell(row=1, column=11).value="Prepared by:"
		F10TB.cell(row=1, column=11).font=font
		F10TB.cell(row=1, column=12).value=preparedBy1
		F10TB.cell(row=1, column=12).font=font

		F10TB.cell(row=2, column=11).value="Date:"
		F10TB.cell(row=2, column=11).font=font
		# F10TB.cell(row=2, column=12).value=datetime.datetime.now()
		F10TB.cell(row=2, column=12).number_format="mm/dd/yyyy"
		F10TB.cell(row=2, column=12).alignment=Alignment(horizontal='left')

		F10TB.cell(row=3, column=11).value="Ref:"
		F10TB.cell(row=3, column=11).font=font
		F10TB.cell(row=3, column=12).value="F10"
		F10TB.cell(row=3, column=12).font=fontRed

		for i in range(1,4):
			F10TB.cell(row=i, column=11).alignment=Alignment(horizontal='right')

		F10TB.cell(row=4, column=2).value="Trial Balance"
		F10TB.cell(row=4, column=2).font=font

		F10TB.cell(row=6, column=1).value="Work done:"
		F10TB.cell(row=6, column=1).font=font


		F10TB.cell(row=8, column=1).value="(to be adjusted for P&L variation; e.g. if YE is different of 31.12)"
		F10TB.cell(row=8, column=1).font=Font(name='Tahoma', size=8, italic=True)




		F10TB.cell(row=14, column=1).value="Class"
		F10TB.cell(row=14, column=2).value="Synt 1"
		F10TB.cell(row=14, column=3).value="Synt 3"
		F10TB.cell(row=14, column=4).value="Synt 4"
		F10TB.cell(row=14, column=5).value="Account"
		F10TB.cell(row=14, column=6).value="Description"
		F10TB.cell(row=14, column=7).value="OB"
		F10TB.cell(row=14, column=8).value="DM"
		F10TB.cell(row=14, column=9).value="CM"
		F10TB.cell(row=14, column=10).value="CB"
		F10TB.cell(row=14, column=11).value="Check"
		F10TB.cell(row=14, column=13).value="Abs CB-OB"
		F10TB.cell(row=14, column=14).value="VAR %"

		# F10TB.cell(row=14, column=16).value="OMF Row"
		# F10TB.cell(row=14, column=17).value="OMF Description"
		# F10TB.cell(row=14, column=18).value="LS"

		F10TB.cell(row=14, column=16).value="Check OB"



		for i in range(1,12):
			F10TB.cell(row=14, column=i).font=font2
			F10TB.cell(row=14, column=i).fill=blueFill
			F10TB.cell(row=14, column=i).border=doubleborder
			F10TB.cell(row=14, column=i).alignment=Alignment(horizontal='left')

		for i in range(13,15):
			F10TB.cell(row=14, column=i).font=font2
			F10TB.cell(row=14, column=i).fill=blueFill
			F10TB.cell(row=14, column=i).border=doubleborder
			F10TB.cell(row=14, column=i).alignment=Alignment(horizontal='left')

		# F10TB.cell(row=14, column=16).font=font2
		# F10TB.cell(row=14, column=16).fill=blueFill
		# F10TB.cell(row=14, column=16).border=doubleborder
		# F10TB.cell(row=14, column=16).alignment=Alignment(horizontal='left')

		# F10TB.cell(row=14, column=17).font=font2
		# F10TB.cell(row=14, column=17).fill=blueFill
		# F10TB.cell(row=14, column=17).border=doubleborder
		# F10TB.cell(row=14, column=17).alignment=Alignment(horizontal='left')

		# F10TB.cell(row=14, column=18).font=font2
		# F10TB.cell(row=14, column=18).fill=blueFill
		# F10TB.cell(row=14, column=18).border=doubleborder
		F10TB.cell(row=14, column=18).alignment=Alignment(horizontal='left')

		F10TB.cell(row=8, column=6).value="Check BS"
		F10TB.cell(row=9, column=6).value="Revenues"
		F10TB.cell(row=10, column=6).value="Expenses"
		F10TB.cell(row=11, column=6).value="Result"
		F10TB.cell(row=11, column=6).border=doubleborder
		F10TB.cell(row=12, column=6).value="Check"

		for i in range (8,13):
			F10TB.cell(row=i, column=6).font=font
			F10TB.cell(row=i, column=6).alignment=Alignment(horizontal='right')

		F10TB.cell(row=8, column=7).value = '=SUMIF(A:A,"BS",G:G)'
		F10TB.cell(row=11, column=7).border=doubleborder

		for i in range (8,14):
			F10TB.cell(row=i, column=7).number_format='#,##0_);(#,##0)'


		F10TB.cell(row=8, column=10).value = '=SUMIF(A:A,"BS",J:J)'
		F10TB.cell(row=9, column=10).value='=SUMIF(B:B,"7",J:J)'
		F10TB.cell(row=10, column=10).value='=SUMIF(B:B,"6",J:J)'
		F10TB.cell(row=11, column=10).value='=SUMIF(C:C,"121",J:J)'
		F10TB.cell(row=11, column=10).border=doubleborder
		F10TB.cell(row=12, column=10).value="=SUM(J9:J10)-J11"
		F10TB.cell(row=12, column=10).font=fontRed

		F10TB.cell(row=8, column=7).value = '=SUMIF(A:A,"BS",G:G)'
		F10TB.cell(row=9, column=7).value='=SUMIF(B:B,"7",G:G)'
		F10TB.cell(row=10, column=7).value='=SUMIF(B:B,"6",G:G)'
		F10TB.cell(row=11, column=7).value='=SUMIF(C:C,"121",G:G)'
		F10TB.cell(row=11, column=7).border=doubleborder
		F10TB.cell(row=12, column=7).value="=SUM(G9:G10)-G11"
		F10TB.cell(row=12, column=7).font=fontRed



		F10TB.cell(row=13, column=12).value="=SUM(K:K)"
		F10TB.cell(row=13, column=12).font=fontRed

		F10TB.cell(row=13,column=11).value="Total diff:"
		F10TB.cell(row=13,column=11).alignment=Alignment(horizontal='right')

		for i in range (8,13):
			F10TB.cell(row=i, column=10).number_format='#,##0_);(#,##0)'



		for i in range (1,10):
			F10TB.cell(row=i, column=15).number_format='#,##0_);(#,##0)'

		for i in range(14,16):
			F10TB.cell(row=1, column=i).font=font2
			F10TB.cell(row=1, column=i).fill=blueFill
			F10TB.cell(row=1, column=i).border=doubleborder
			F10TB.cell(row=1, column=i).alignment=Alignment(horizontal='left')

		"Importing Data"

		if isChecked4=="":
			try:

				acc=Account+Accountp

				mylist2 = list(dict.fromkeys(acc))
				mylist=[]
				for xxx in range(0,len(mylist2)):
					mylist.append(str(mylist2[xxx]))
				mylist.sort()

				print(mylist)
			except:
				# acc=Account
				mylist2 = list(dict.fromkeys(Account))
				mylist=[]
				for xxx in range(0,len(mylist2)):
					mylist.append(str(mylist2[xxx]))
				mylist.sort()

				print(mylist)
		else:
			mylist2 = list(dict.fromkeys(Account))
			mylist=[]
			for xxx in range(0,len(mylist2)):
				mylist.append(str(mylist2[xxx]))
			mylist.sort()

			print(mylist)


		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=5).value=int(mylist[i-1])

		for i in range  (  1, len(mylist)+1):
			if(mylist[i-1] in Account):
				F10TB.cell(row=i+14, column=6).value='=VLOOKUP(E{0},PBC_CY!D:E,2,0)'.format(i+14)
			else:
				F10TB.cell(row=i+14, column=6).value='=VLOOKUP(E{0},PBC_PY!D:E,2,0)'.format(i+14)

		for i in range (1,len(mylist)+1):
			x=str(mylist[i-1])
			y=str(x[:4])
			F10TB.cell(row=i+14, column=4).value=str(y)

		for i in range (1,len(mylist)+1):
			x=str(mylist[i-1])
			y=x[:3]
			F10TB.cell(row=i+14, column=3).value=str(y)

		for i in range (1,len(mylist)+1):
			F10TB.cell(row=i+14, column=2).value='=Left(E{0},1)'.format(i+14)


		for i in range(1, len(mylist)+1):
				F10TB.cell(row=i+14, column=1).value='=IF(B{0}<"6","BS",IF(AND(B{0}>"5",B{0}<"8"),"PL","Other Account-Off TB"))'.format(i+14)
		 

		"Calculation"

		for i in range(1, len(mylist)+1):
			if(mylist[i-1] in Account):

				if(int(str(mylist[i-1])[:1])<6):
					F10TB.cell(row=i+14,column=7).value='=SUMIF(PBC_CY!D:D,E{0},PBC_CY!F:F)-SUMIF(PBC_CY!D:D,E{0},PBC_cY!G:G)'.format(i+14)
					F10TB.cell(row=i+14,column=7).number_format='#,##0_);(#,##0)'
			else:
				F10TB.cell(row=i+14,column=7).value='=SUMIF(PBC_PY!D:D,E{0},PBC_PY!F:F)'.format(i+14)
				F10TB.cell(row=i+14,column=7).number_format='#,##0_);(#,##0)'
			if(int(str(mylist[i-1])[:1])==6):
				F10TB.cell(row=i+14,column=7).value='=SUMIF(PBC_PY!D:D,E{0},PBC_PY!F:F)'.format(i+14)
				F10TB.cell(row=i+14,column=7).number_format='#,##0_);(#,##0)'
			if(int(str(mylist[i-1])[:1])==7):
				F10TB.cell(row=i+14,column=7).value='=SUMIF(PBC_PY!D:D,E{0},PBC_PY!F:F)'.format(i+14)
				F10TB.cell(row=i+14,column=7).number_format='#,##0_);(#,##0)'

		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=8).value='=SUMIF(PBC_CY!D:D,E{0},PBC_CY!H:H)'.format(i+14)
			F10TB.cell(row=i+14,column=8).number_format='#,##0_);(#,##0)'

		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=9).value='=SUMIF(PBC_CY!D:D,E{0},PBC_CY!I:I)'.format(i+14)
			F10TB.cell(row=i+14,column=9).number_format='#,##0_);(#,##0)'

		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=10).value='=IF(B{0}<"6",SUMIF(PBC_CY!D:D,E{0},PBC_CY!J:J)-SUMIF(PBC_CY!D:D,E{0},PBC_CY!K:K),IF(AND(B{0}="6",C{0}="609",H{0}>0),-H{0},IF(AND(B{0}="6",H{0}<>I{0}),H{0}-I{0},IF(B{0}="6",H{0},IF(AND(B{0}="7",C{0}="709",I{0}<0),I{0},IF(AND(B{0}="7",C{0}="711"),-$U$6,IF(AND(B{0}="7",C{0}="712"),-$U$8,IF(AND(B{0}="7",H{0}<>I{0}),H{0}-I{0},IF(AND(B{0}="7",I{0}>0),-I{0},IF(AND(B{0}="7",I{0}<0),I{0},0))))))))))'.format(i+14)
			F10TB.cell(row=i+14,column=10).number_format='#,##0_);(#,##0)'


		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column =11).value='=IF(A{0}="BS",G{0}+H{0}-I{0}-J{0},IF(AND(A{0}="PL",H{0}<>I{0}),H{0}-I{0}-J{0},H{0}-I{0}))'.format(i+14)
			F10TB.cell(row=i+14,column=11).number_format='#,##0_);(#,##0)'
			F10TB.cell(row=i+14, column=11).font=fontRed

		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=13).value='=IF(A{0}="BS",J{0}-G{0},"")'.format(i+14)
			F10TB.cell(row=i+14,column=13).number_format='#,##0_);(#,##0)'

		F10TB.cell(row=13, column=12).number_format='#,##0_);(#,##0)'

		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14, column=14).value='=IF(A{0}="BS",IF(AND(G{0}=0,J{0}=0),0,IF(AND(G{0}=0,J{0}>0),1,IF(AND(G{0}=0,J{0}<0),-1,IF(AND(J{0}=0,G{0}>0),-1,IF(AND(J{0}=0,G{0}<0),1,J{0}/G{0}-1))))),"")'.format(i+14)
			F10TB.cell(row=i+14,column=14).number_format="0.0%"

		# for i in range(1, len(mylist)+1):
		# 	x=str(mylist[i-1])
		# 	F10TB.cell(row=i+14,column=16).value='=if('+x[:1]+'<6'+",iferror(vlookup(D{0},'BS Mapping'!A:C,3,0),vlookup(C{0},'BS Mapping'!A:C,3,0)),iferror(vlookup(D{0},'PL Mapping'!A:C,3,0),vlookup(C{0},'PL Mapping'!A:C,3,0)))".format(i+14)

		# for i in range(1, len(mylist)+1):
		# 	x=str(mylist[i-1])
		# 	F10TB.cell(row=i+14,column=17).value='=if('+x[:1]+'<6'+",iferror(vlookup(D{0},'BS Mapping'!A:D,4,0),vlookup(C{0},'BS Mapping'!A:D,4,0)),iferror(vlookup(D{0},'PL Mapping'!A:D,4,0),vlookup(C{0},'PL Mapping'!A:D,4,0)))".format(i+14)
		# for i in range(1, len(mylist)+1):
		# 	x=str(mylist[i-1])
		# 	F10TB.cell(row=i+14,column=18).value='=if('+x[:1]+'<6'+",iferror(vlookup(D{0},'BS Mapping'!A:E,5,0),vlookup(C{0},'BS Mapping'!A:E,5,0)),iferror(vlookup(D{0},'PL Mapping'!A:E,5,0),vlookup(C{0},'PL Mapping'!A:E,5,0)))".format(i+14)
		for i in range(1, len(mylist)+1):
			F10TB.cell(row=i+14,column=16).value="=G{0}-SUMIF(PBC_PY!D:D,E{0},PBC_PY!F:F)".format(i+14)
			F10TB.cell(row=i+14,column=16).number_format='#,##0_);(#,##0)'
			F10TB.cell(row=i+14,column=16).font=fontRed
			# F10TB.cell(row=i+14,column=16).value='=if('+x[:1]+'<6,iferror(vlookup('+str(x[0:4])+",'BS Mapping std'!A:E,5,0),vlookup("+str(x[0:3])+",'BS Mapping std'!A:E,5,0)),iferror(vlookup("+str(x[0:4])+",'PL mapping Std'!A:E,5,0),vlookup("+str(x[0:3])+",'PL mapping Std'!A:E,5,0))"
		"Closing 711"

		F10TB.cell(row=1, column=14).value="Closing 711"
		F10TB.cell(row=1, column=14).font=font2
		F10TB.cell(row=1, column=14).alignment=Alignment(horizontal='right')

		F10TB.cell(row=1, column=14).value="Acc."
		F10TB.cell(row=1, column=15).value="OB"
		F10TB.cell(row=1, column=16).value="CB"
		F10TB.cell(row=1, column=17).value="VAR"

		F10TB.cell(row=2, column=14).value="331"
		F10TB.cell(row=3, column=14).value="341"
		F10TB.cell(row=4, column=14).value="345"
		F10TB.cell(row=5, column=14).value="348"
		F10TB.cell(row=6, column=14).value="Total:"
		F10TB.cell(row=6, column=14).font=font
		F10TB.cell(row=6, column=14).alignment=Alignment(horizontal='right' )

		F10TB.cell(row=8, column=14).value="332" 
		F10TB.cell(row=8, column=14).font=font

		for i in range(14,18):
			F10TB.cell(row=1, column=i).font=font2
			F10TB.cell(row=1, column=i).fill=blueFill
			F10TB.cell(row=1, column=i).border=doubleborder
			F10TB.cell(row=1, column=i).alignment=Alignment(horizontal='left')

		for i in range (2,9):
			F10TB.cell(row=i, column=14).font=font
			F10TB.cell(row=i, column=14).alignment=Alignment(horizontal='right')

		for i in range (14,18):
			F10TB.cell(row=5, column=i).border=doubleborder

		for i in range (2, 6):
			F10TB.cell(row=i, column=15).value='=SUMIF(C:C,N{0},G:G)'.format(i)
			F10TB.cell(row=i, column=16).value='=SUMIF(C:C,N{0},J:J)'.format(i)
			F10TB.cell(row=i, column=17).value='=P{0}-O{0}'.format(i)

		F10TB.cell(row=6, column=15).value='=SUM(O2:O5)'
		F10TB.cell(row=6, column=16).value='=SUM(P2:P5)'
		F10TB.cell(row=6, column=17).value='=P6-O6'
		F10TB.cell(row=6, column=17).font=font2
		F10TB.cell(row=6, column=17).fill=blueFill

		for i in range (15,18):
			F10TB.cell(row=6, column=i).font=font2
		F10TB.cell(row=14,column=16).fill=blueFill
		F10TB.cell(row=14,column=16).font=font2

		for i in range(2,14):
			for j in range(15,18):  
				F10TB.cell(row=i, column=j).number_format='#,##0_);(#,##0)'

		F10TB.cell(row=7, column=14).value="Closing 712"
		F10TB.cell(row=7, column=14).font=font2
		F10TB.cell(row=7, column=14).alignment=Alignment(horizontal='right')

		F10TB.cell(row=8, column=15).value='=SUMIF(C:C,N8,G:G)'
		F10TB.cell(row=8, column=16).value='=SUMIF(C:C,N8,J:J)'
		F10TB.cell(row=8, column=17).value='=P8-O8'
		F10TB.cell(row=8, column=17).fill=blueFill
		F10TB.cell(row=8, column=17).font=font2

		for i in range (6,10):
			F10TB.cell(row=11, column=i).border=doubleborder

		F10TB.auto_filter.ref = 'A14:P14'

		c = F10TB['B15']
		F10TB.freeze_panes = c


		x
		"Adjust Column Width"

		for col in F10TB.columns:
			max_length = 0
			for cell in col:
				if cell.coordinate in F10TB.merged_cells:
					continue
				try:
					if len(str(cell.value)) > max_length:
						max_length = len(cell.value)
				except:
					pass
			adjusted_width=(max_length-5)


		listanoua=['F','G','H','I','J','K','M','N','O','L']
		for column in ascii_uppercase:
			for i in listanoua:
				if (column==i):
					F10TB.column_dimensions[column].width =15

		listanoua2=['A']
		for column in ascii_uppercase:
			for i in listanoua2:
				if (column==i):
					F10TB.column_dimensions[column].width = 10


		
		file_path=os.path.join(folderpath, "F100 Trial Balance.xlsx")
		myorder=[3,2,1]
		output._sheets =[output._sheets[i] for i in myorder]
		output.save(folderpath+"/Trial Balance.xlsx")
		return send_from_directory(folderpath,"Trial Balance.xlsx",as_attachment=True)

		# print(text)


if __name__ == '__main__':
   	app.run()
