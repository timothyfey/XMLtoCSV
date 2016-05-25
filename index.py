#!/usr/bin/python
#------------------------------------------
# XML to CSV converter to be used for billing portal
#     TEF   -   Original creation 8/18/2015



print '''Content-Type: text/html

<html>
<head><title>***XML to CSV Converter***</title></head>
<body>
<img src="WestparkSmallLogo.jpg"/><br><br>
<font size="5"><b>Billing XML to CSV Converter</b></font><br><br>
<a href="invoice-in/">Click here for previous invoice uploads</a><br><br>
<a href="invoice-out/">Click here for previous invoice conversions</a><br><br><br>
<form action = '' method = "post" enctype="multipart/form-data">
File: <input type="file" name="file"><br><br><br>
<input type=submit name=submit value="Run" ></body>
</form></html>'''

import cgi
import time
import codecs
form = cgi.FieldStorage()

#Importing datetime to be used in generated file name
import datetime
mylist = []
present = datetime.date.today()
present2 = datetime.datetime.now().strftime('%Y-%m-%d__%H:%M:%S')
mylist.append(present)

#If condition will trigger when the "Run" button is clicked.
if str(form.getvalue("submit"))!="None":
	
	
	
	
	FileItem = form["file"]
	from xml.etree import ElementTree
	import csv

	#Since I can't work directly from the file that is uploaded by the user, I create a new file locally
	#that is the exact same as the uploaded file, name and contents.	
	fb = open('invoice-in/' + FileItem.filename,'w')
	fb.write(form.getvalue("file"))
	fb.close()
	
	
	with open('invoice-in/' + FileItem.filename, 'rt') as f:
		
		tree = ElementTree.parse(f)
		root = tree.getroot()


	headers=[]
	FormattedNames=["Invoice ID","Company","Invoice Date","Total","PO Number","Notes","Line Item Desc","Line Item Rate","Line Item Quantity","Line Item Percentage","Line Item Total"]

	i=0
	k=0
	for invoice in root.findall('INVOICE'):
		j=0
		for tag in invoice.getiterator():
			j+=1
			if i < 1 or j > k:
				if str(tag.tag) not in headers:
					headers.append(str(tag.tag))            
		i+=1
		if k < j:
			k=j
	fileoutname = "CSV_Billing_" + present2+  ".csv"
			
	with codecs.open ('invoice-out/' + fileoutname, 'wb', encoding='latin-1') as myfile:
		wr=csv.writer(myfile, quoting=csv.QUOTE_ALL)
		wr.writerow(FormattedNames)
		
	   
		for invoice in root.findall('INVOICE'):
			atts = []
			atts2=[]
			inner = []            
			Invoice_Number=[]
			Invdet_Descr=[]
			Customer_Name=[]
			Invoice_Date=[]
			Invoice_PO_Number=[]
			Invoice_Due_Date=[]
			Invdet_Tax1=[]
			Invdet_Tax2=[]
			Invdet_TaxTotal=[]
			Invoice_Account_Total_Due=[]
			Customer_Account=[]
			Invdet_Price_Each=[]
			Invdet_Qty=[]
			Invdet_Charge=[]
			DUNNING_MSG=[]
			
			TaxTotal = 0.00
			amt1=""       
			
			for x in headers:
				
				for tag in invoice.getiterator():
					
					if str(tag.tag)=="Invoice_Number" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Invoice_Number.append(str(tag.text))
					if str(tag.tag)=="Invdet_Descr" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Invdet_Descr.append(str(tag.text))
					if str(tag.tag)=="Customer_Name" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Customer_Name.append(str(tag.text))
					if str(tag.tag)=="Customer_Account" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Customer_Account.append(str(tag.text))
					if str(tag.tag)=="Invoice_Date" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Invoice_Date.append(str(tag.text))
					if str(tag.tag)=="Invoice_PO_Number" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Invoice_PO_Number.append(str(tag.text))
					if str(tag.tag)=="Invoice_Due_Date" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							Invoice_Due_Date.append(str(tag.text))
					#if str(tag.tag)=="Invdet_Tax1" and str(tag.tag)==str(x):               
					   # if str(tag.text)!="None":
						  #  Invdet_Tax1.append(str(tag.text))
					if str(tag.tag)=="Invdet_Tax2" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							if str(tag.text) == "TX0":
								Invdet_Tax2.append("0")
							elif str(tag.text) == "TX1":
								Invdet_Tax2.append("8.25")
							else:
								Invdet_Tax2.append("")

					if str(tag.tag)=="Invoice_Amount" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text)
							amt1=amt1.replace("$","")
							amt1=amt1.replace(",","")
							
							Invoice_Account_Total_Due.append(amt1)
					if str(tag.tag)=="Invdet_Price_Each" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text)
							amt1=amt1.replace("$","")
							amt1=amt1.replace(",","")
							if "(" in amt1:
								amt1=amt1.replace("(","-")
								amt1=amt1.replace(")","")
							Invdet_Price_Each.append(amt1)
							
					if str(tag.tag)=="Invdet_Qty" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							if str(tag.text)=="0":
								Invdet_Qty.append(" ")
							else:
								Invdet_Qty.append(str(tag.text))
						
					if str(tag.tag)=="Invdet_Charge" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text)
							amt1=amt1.replace("$","")
							amt1=amt1.replace(",","")
							if amt1 == "0.00":
								amt1=" "
							if "(" in amt1:
								amt1=amt1.replace("(","-")
								amt1=amt1.replace(")","")
							Invdet_Charge.append(amt1)
					if str(tag.tag)=="DUNNING_MSG" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							DUNNING_MSG.append(str(tag.text))

					if str(tag.tag)=="Invdet_Tax1Amt" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text)
							amt1=amt1.replace("$","")
							amt1=amt1.replace(",","")
							if "(" in amt1:
								amt1=amt1.replace("(","")
								amt1=amt1.replace(")","")
								amt1=float(amt1)
								TaxTotal = TaxTotal - amt1
							else:
								amt1=float(amt1)
								TaxTotal = TaxTotal + amt1
					if str(tag.tag)=="Invdet_Tax2Amt" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text).replace("$","")
							amt1=amt1.replace(",","")
							if "(" in amt1:
								amt1=amt1.replace("(","")
								amt1=amt1.replace(")","")
								amt1=float(amt1)
								TaxTotal = TaxTotal - amt1
							else:
								amt1=float(amt1)
								TaxTotal = TaxTotal + amt1
					if str(tag.tag)=="Invdet_Tax3Amt" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text).replace("$","")
							amt1=amt1.replace(",","")
							if "(" in amt1:
								amt1=amt1.replace("(","")
								amt1=amt1.replace(")","")
								amt1=float(amt1)
								TaxTotal = TaxTotal - amt1
							else:
								amt1=float(amt1)
								TaxTotal = TaxTotal + amt1
					if str(tag.tag)=="Invdet_Tax4Amt" and str(tag.tag)==str(x):               
						if str(tag.text)!="None":
							amt1=str(tag.text).replace("$","")
							amt1=amt1.replace(",","")
							if "(" in amt1:
								amt1=amt1.replace("(","")
								amt1=amt1.replace(")","")
								amt1=float(amt1)
								TaxTotal = TaxTotal - amt1
							else:
								amt1=float(amt1)
								TaxTotal = TaxTotal + amt1                        
							
			
			name1=Customer_Account[0]
			if name1[0]=='0':
				name1=name1[1:]
			# Removed Hyphen (but kept space) as this was causing problems with import -JMP
			#name = name1 + " \xE2\x80\x93 " + Customer_Name[0]
			#name.decode("Windows-1252")
			name = name1 + " " + Customer_Name[0]
			
			stepthru=0
			temp1 =""
			xtwo=[]
			for item in Invdet_Charge:				
				if item != "0.00":
					
					if Invdet_Price_Each[stepthru] == "0.00":
						#xtwo.append(Invdet_Price_Each[stepthru])
						Invdet_Price_Each.pop(stepthru)
						Invdet_Price_Each.insert(stepthru, item)
						Invdet_Qty.pop(stepthru)
						Invdet_Qty.insert(stepthru, "1")
						
				stepthru = stepthru + 1	

				
			if len(Invoice_Number)>0:     
				atts.append(','.join(Invoice_Number))           
			if len(Customer_Name)>0 :       
				atts.append(name)
			if len(Invoice_Date)>0  :     
				atts.append(','.join(Invoice_Date))
			#if len(Invoice_Due_Date)>0  : 
				#atts.append(','.join(Invoice_Due_Date))
			#atts.append(','.join(Invdet_Tax1))
			#atts.append(','.join(Invdet_Tax2))
			#atts.append(str(TaxTotal))
			atts.append(','.join(Invoice_Account_Total_Due))   
			atts.append(','.join(Invoice_PO_Number))
			atts.append(','.join(DUNNING_MSG))
			if len(Invdet_Descr)>0 :         
				atts.append(','.join(Invdet_Descr))
			atts.append(','.join(Invdet_Price_Each))
			atts.append(','.join(Invdet_Qty))
			atts.append(" ")
			#Invdet_Charge[1]=Str("TEST")
			atts.append(','.join(Invdet_Charge)) 
			atts.append(','.join(xtwo))  			
				
			if len(atts)>0:
				wr.writerow(atts)
			

		string = '<a href = "invoice-out/' + fileoutname + '"> Click here for your file!! </a>'
		print string
	
	
	




