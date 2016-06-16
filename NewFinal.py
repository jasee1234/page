#!/usr/bin/env python


from xlrd import open_workbook
import sys
import urllib.request
import csv
from xlsxwriter import Workbook


class Origincharges(object):
	charges_20=[]
	charges_40=[]
	charges_40HC=[]
	charges_45HC=[]
	filled_20=False
	filled_40=False
	filled_40HC=False
	filled_45HC=False
	def __init__(self, portname,charges_20,charges_40,charges_40HC,charges_45HC):
		self.portname = portname
		self.charges_20 = charges_20
		self.charges_40= charges_40
		self.charges_40HC=charges_40HC
		self.charges_45HC=charges_45HC
	def send(self):
		print(self.portname,self.charges_20,self.charges_40,self.charges_40HC,self.charges_45HC)

	def addcharges_20(self,charges_20):
		self.charges_20=charges_20	
	def addcharges_40(self,charges_40):
		self.charges_40=charges_40
	def addcharges_40HC(self,charges_40HC):
		self.charges_40HC=charges_40HC	
	def addcharges_45HC(self,charges_45HC):
		self.charges_45HC=charges_45HC			


class Destinationcharges(object):
	charges_20=[]
	charges_40=[]
	charges_40HC=[]
	charges_45HC=[]
	filled_20=False
	filled_40=False
	filled_40HC=False
	filled_45HC=False
	ddf=''
	def __init__(self, portname,charges_20,charges_40,charges_40HC,charges_45HC):
		self.portname = portname
		self.charges_20 = charges_20
		self.charges_40= charges_40
		self.charges_40HC=charges_40HC
		self.charges_45HC=charges_45HC
	def send(self):
		print(self.portname,self.ddf)

	def addcharges_20(self,charges_20):
		self.charges_20=charges_20	
	def addcharges_40(self,charges_40):
		self.charges_40=charges_40
	def addcharges_40HC(self,charges_40HC):
		self.charges_40HC=charges_40HC	
	def addcharges_45HC(self,charges_45HC):
		self.charges_45HC=charges_45HC
	def conversion(self):
		for i in range(len(self.charges_20)):
			if str(self.charges_20[i][0])=='DDF' and self.filled_20:
				self.ddf=str(self.charges_20[i][2])[-3:]
		for i in range(len(self.charges_40)):
			if str(self.charges_40[i][0])=='DDF' and self.filled_40:
				self.ddf=str(self.charges_40[i][2])[-3:]
		for i in range(len(self.charges_40HC)):
			if str(self.charges_40HC[i][0])=='DDF' and self.filled_40HC:
				self.ddf=str(self.charges_40HC[i][2])[-3:]
		for i in range(len(self.charges_45HC)):
			if str(self.charges_45HC[i][0])=='DDF' and self.filled_45HC:
				self.ddf=str(self.charges_45HC[i][2])[-3:]	


class Oceancharges(object):
	charges_20=[]
	charges_40=[]
	charges_40HC=[]
	charges_45HC=[]
	filled_20=False
	filled_40=False
	filled_40HC=False
	filled_45HC=False
	def __init__(self, fromport,toport,charges_20,charges_40,charges_40HC,charges_45HC):
		self.fromport = fromport
		self.toport=toport
		self.charges_20 = charges_20
		self.charges_40= charges_40
		self.charges_40HC=charges_40HC
		self.charges_45HC=charges_45HC
	def send(self):
		print(self.fromport,self.toport,self.charges_40,charges_40HC,charges_45HC)
	def addcharges_20(self,charges_20):
		self.charges_20=charges_20	
	def addcharges_40(self,charges_40):
		self.charges_40=charges_40
	def addcharges_40HC(self,charges_40HC):
		self.charges_40HC=charges_40HC	
	def addcharges_45HC(self,charges_45HC):
		self.charges_45HC=charges_45HC


mapping={}
description={}
Total=[]
book = open_workbook(sys.argv[1])
abbr=book.sheet_by_index(2)
port=book.sheet_by_index(1)

print('Processing the files')

#Copying the Abbreviations from Legend

for row in range(2,45):  # 2,45 are the numbers from charge code,upload code start and end
		prices=[]
		for col in range(0,3):
			prices.append((abbr.cell(row,col).value))
		Total.append(prices)

#print(Total)

for i in range(len(Total)):
	mapping[Total[i][0]]=Total[i][2]
	description[Total[i][0]]=Total[i][1]

#print(mapping)

# Copying the port codes

Portcodes={}
for row in range(1,port.nrows):  
			Portcodes[port.cell(row,0).value]=port.cell(row,2).value
#print (Portcodes)

book = open_workbook(sys.argv[2])
port=book.sheet_by_index(0)
for row in range(1,port.nrows):  
			portname=str((port.cell(row,2).value).split("(")[0])+' ,'+str((port.cell(row,1).value))[0:2]
			pcode=port.cell(row,1).value
			if portname not in Portcodes:
				Portcodes[portname]=pcode

port=book.sheet_by_index(1)
for row in range(1,port.nrows):  
			portname=(port.cell(row,1).value).split("(")[0]+' ,'+str((port.cell(row,0).value))[0:2]
			pcode=port.cell(row,0).value
			if portname not in Portcodes:
				Portcodes[portname]=pcode

				
#copying the input file to lists
Freight_charges= sys.argv[3]
Indian_origin_charges=sys.argv[4]

for no in range(5,len(sys.argv)):
	filename = sys.argv[no]
	wb = open_workbook(filename)
	Total=[]
	s=wb.sheet_by_index(1)
	for row in range(7,s.nrows):  # 7 from which line it starts 
		prices=[]
		for col in range(0,s.ncols):
			prices.append((s.cell(row,col).value))
		Total.append(prices)
	#print(len(Total))

	for i in range(0,len(Total)):
		if Total[i][8]=='PER_DOC':
			Total[i][8]='PER_BL'	
		if Total[i][7]=='BAS':
			#print(i)
			if Total[i][9]!='':
				Total[i][9]=str(float(str(Total[i][9])[0:-3])+float(Freight_charges))+str(Total[i][9])[-3:]
			if Total[i][10]!='':
				Total[i][10]=str(float(str(Total[i][10])[0:-3])+float(Freight_charges))+str(Total[i][10])[-3:]
			if Total[i][11]!='':		
				Total[i][11]=str(float(str(Total[i][11])[0:-3])+float(Freight_charges))+str(Total[i][11])[-3:]
			if Total[i][12]!='':
				Total[i][12]=str(float(str(Total[i][12])[0:-3])+float(Freight_charges))+str(Total[i][12])[-3:]
			
	Currency=[]

	for i in range(len(Total)):
		if str(Total[i][9])[-3:] not in Currency:
			Currency.append(str(Total[i][9])[-3:])
	#print(Currency)

	conv=[]
	for i in range(0,len(Currency)):
		conv.append(Currency[i]+str('USD'))
	#print(conv)
	x=''
	for i in range(len(conv)):
		x=x+conv[i]+'=X+'
	s=x[0:-1]
	f='nl1d1t1'
	#print(x[0:-1])
	url = "http://download.finance.yahoo.com/d/quotes.csv?s=%s&f=%s" % (s,f)
	#print(url)
	info = urllib.request.urlopen(url)
	with open('info.csv','wb') as output:
	  output.write(info.read())


	conversion={}
	conversion['USD/USD']=float('1')

	with open('info.csv', newline='\n') as f:
		reader = csv.reader(f, delimiter=',')
		for row in reader:
			try:	
				conversion[row[0]]=float(row[1])
			except:
				print('')

	for i in range(len(Total)):
		if str(Total[i][0])[-2:]=='IN' and str(Total[i][7])=='ODF':
			if Total[i][9]!='':
				Total[i][9]=str(float(str(Total[i][9])[0:-3])+float(Indian_origin_charges))+str(Total[i][9])[-3:]
			if Total[i][10]!='':
				Total[i][10]=str(float(str(Total[i][10])[0:-3])+float(Indian_origin_charges))+str(Total[i][10])[-3:]
			if Total[i][11]!='':		
				Total[i][11]=str(float(str(Total[i][11])[0:-3])+float(Indian_origin_charges))+str(Total[i][11])[-3:]
			if Total[i][12]!='':
				Total[i][12]=str(float(str(Total[i][12])[0:-3])+float(Indian_origin_charges))+str(Total[i][12])[-3:]
			


	Origin_list=[]
	Origin_map={}

	Indian=[]

	for i in range(len(Total)):
		if str(Total[i][0])[-2:]=='IN':
			Indian.append(Total[i])
		if str(Total[i][1])[-2:]=='IN':
			Indian.append(Total[i])

	Total=[]
	Total=Indian

	index=0
	i=0
	while i<(len(Total))-1:
		for j in range(i,len(Total)):
			if Total[j][1]!=Total[i][1] or Total[j][0]!=Total[i][0] or j==len(Total)-1:
				m=j
				break
		charges_20=[]
		charges_40=[]
		charges_40HC=[]
		charges_45HC=[]
		for k in range(i,m):
			if Total[k][9]!='' and mapping[Total[k][7]]=='origin':	
				charges_20.append([Total[k][7],Total[k][8],Total[k][9]])
			if Total[k][10]!='' and  mapping[Total[k][7]]=='origin':	
				charges_40.append([Total[k][7],Total[k][8],Total[k][10]])
			if Total[k][11]!='' and mapping[Total[k][7]]=='origin':	
				charges_40HC.append([Total[k][7],Total[k][8],Total[k][11]])
			if Total[k][12]!='' and  mapping[Total[k][7]]=='origin':	
				charges_45HC.append([Total[k][7],Total[k][8],Total[k][12]])

		if Total[i][0]  not in Origin_map:
			Origin_map[Total[i][0]]=index
			Origin_list.append(Origincharges(Total[i][0],charges_20,charges_40,charges_40HC,charges_45HC))
			if len(charges_20)!=0:
				Origin_list[index].filled_20=True
			if len(charges_40)!=0:
				Origin_list[index].filled_40=True
			if len(charges_40HC)!=0:
				Origin_list[index].filled_40HC=True
			if len(charges_45HC)!=0:
				Origin_list[index].filled_45HC=True
			index=index+1	
		else:
			l=Origin_map[Total[i][0]]
			if Origin_list[l].filled_20==False and len(charges_20)!=0:
				Origin_list[l].addcharges_20(charges_20)
				Origin_list[l].filled_20=True
			if Origin_list[l].filled_40==False and len(charges_40)!=0:
				Origin_list[l].addcharges_40(charges_40)
				Origin_list[l].filled_40==True
			if Origin_list[l].filled_40HC==False and len(charges_40HC)!=0:
				Origin_list[l].addcharges_40HC(charges_40HC)
				Origin_list[l].filled_40HC=True
			if Origin_list[l].filled_45HC==False and len(charges_45HC)!=0:
				Origin_list[l].addcharges_45HC(charges_45HC)
				Origin_list[l].filled_45HC=True

		i=m

	print('Creating the Origin charges file')
	output = open('/var/www/html/Info/Maersk/Origincharges.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container-type'+','+'Container size'+','+'Unit'+','+'Rate/unit'+'\n')
	for i in range(0,len(Origin_list)):
		if Origin_list[i].portname in Portcodes:
			output.write('\n'+'port'+','+str(Portcodes[Origin_list[i].portname])+'\n')
		else:
			output.write('\n'+'port'+','+str(Origin_list[i].portname).replace(","," ")+'\n')

		for j in range(0,len(Origin_list[i].charges_20)):
			output.write(str(description[str(Origin_list[i].charges_20[j][0])])+','+str(Origin_list[i].charges_20[j][2])[-3:]+','+'standard'+','+'20,'+str(Origin_list[i].charges_20[j][1])+','+str(Origin_list[i].charges_20[j][2])[0:-3]+'\n')

		for j in range(0,len(Origin_list[i].charges_40)):
			output.write(str(description[str(Origin_list[i].charges_40[j][0])])+','+str(Origin_list[i].charges_40[j][2])[-3:]+','+'standard'+','+'40,'+str(Origin_list[i].charges_40[j][1])+','+str(Origin_list[i].charges_40[j][2])[0:-3]+'\n')

		for j in range(0,len(Origin_list[i].charges_40HC)):
			output.write(str(description[str(Origin_list[i].charges_40HC[j][0])])+','+str(Origin_list[i].charges_40HC[j][2])[-3:]+','+'standard'+','+'40HC,'+str(Origin_list[i].charges_40HC[j][1])+','+str(Origin_list[i].charges_40HC[j][2])[0:-3]+'\n')

		for j in range(0,len(Origin_list[i].charges_45HC)):
			output.write(str(description[str(Origin_list[i].charges_45HC[j][0])])+','+str(Origin_list[i].charges_45HC[j][2])[-3:]+','+'standard'+','+'45HC,'+str(Origin_list[i].charges_45HC[j][1])+','+str(Origin_list[i].charges_45HC[j][2])[0:-3]+'\n')


	i=0
	index=0
	Destination_list=[]
	Dest_map={}

	while i<(len(Total))-1:
		for j in range(i,len(Total)):
			if Total[j][1]!=Total[i][1] or Total[j][0]!=Total[i][0] or j==len(Total)-1:
				m=j
				break
		charges_20=[]
		charges_40=[]
		charges_40HC=[]
		charges_45HC=[]
		for k in range(i,m):
			if mapping[Total[k][7]]=='Destination':	
				charges_20.append([Total[k][7],Total[k][8],Total[k][9]])
			if Total[k][10]!='' and  mapping[Total[k][7]]=='Destination':	
				charges_40.append([Total[k][7],Total[k][8],Total[k][10]])
			if Total[k][11]!='' and mapping[Total[k][7]]=='Destination':	
				charges_40HC.append([Total[k][7],Total[k][8],Total[k][11]])
			if Total[k][12]!='' and  mapping[Total[k][7]]=='Destination':	
				charges_45HC.append([Total[k][7],Total[k][8],Total[k][12]])


		if Total[i][1] not in Dest_map:
			Dest_map[Total[i][1]]=index
			Destination_list.append(Destinationcharges(Total[i][1],charges_20,charges_40,charges_40HC,charges_45HC))
			if len(charges_20)!=0:
				Destination_list[index].filled_20=True
			if len(charges_40)!=0:
				Destination_list[index].filled_40=True
			if len(charges_40HC)!=0:
				Destination_list[index].filled_40HC=True
			if len(charges_45HC)!=0:
				Destination_list[index].filled_45HC=True
			index=index+1	
		else:
			l=Dest_map[Total[i][1]]
			if Destination_list[l].filled_20==False and len(charges_20)!=0:
				Destination_list[l].addcharges_20(charges_20)
				Destination_list[l].filled_20=True
			if Destination_list[l].filled_40==False and len(charges_40)!=0:
				Destination_list[l].addcharges_40(charges_40)
				Destination_list[l].filled_40==True
			if Destination_list[l].filled_40HC==False and len(charges_40HC)!=0:
				Destination_list[l].addcharges_40HC(charges_40HC)
				Destination_list[l].filled_40HC=True
			if Destination_list[l].filled_45HC==False and len(charges_45HC)!=0:
				Destination_list[l].addcharges_45HC(charges_45HC)
				Destination_list[l].filled_45HC=True

		i=m

	print('Creating the Destination Charges file')

	output = open('/var/www/html/Info/Maersk/Destinationcharges.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container-type'+','+'Container size'+','+'Unit'+','+'Rate/unit'+'\n')
	for i in range(0,len(Destination_list)):
		Destination_list[i].conversion()

		if Destination_list[i].portname in Portcodes:
			output.write('\n'+'port'+','+str(Portcodes[Destination_list[i].portname])+'\n')
		else:
			output.write('\n'+'port'+','+str(Destination_list[i].portname).replace(","," ")+'\n')
		
		for j in range(0,len(Destination_list[i].charges_20)):
			if Destination_list[i].ddf!='':
				factor=conversion[str(Destination_list[i].charges_20[j][2])[-3:]+'/USD']/conversion[Destination_list[i].ddf+'/USD']	
				price=float(str(Destination_list[i].charges_20[j][2])[0:-3])*factor
				price=round(price,2)
				output.write(str(description[str(Destination_list[i].charges_20[j][0])])+','+str(Destination_list[i].ddf)+','+'standard'+','+'20,'+str(Destination_list[i].charges_20[j][1])+','+str(price)+'\n')
			else:		
				output.write(str(description[str(Destination_list[i].charges_20[j][0])])+','+str(Destination_list[i].charges_20[j][2])[-3:]+','+'standard'+','+'20,'+str(Destination_list[i].charges_20[j][1])+','+str(Destination_list[i].charges_20[j][2])[0:-3]+'\n')

		for j in range(0,len(Destination_list[i].charges_40)):
			if Destination_list[i].ddf!='':
				factor=conversion[str(Destination_list[i].charges_40[j][2])[-3:]+'/USD']/conversion[Destination_list[i].ddf+'/USD']	
				price=float(str(Destination_list[i].charges_40[j][2])[0:-3])*factor
				price=round(price,2)
				output.write(str(description[str(Destination_list[i].charges_40[j][0])])+','+str(Destination_list[i].ddf)+','+'standard'+','+'40,'+str(Destination_list[i].charges_40[j][1])+','+str(price)+'\n')
			else:
				output.write(str(description[str(Destination_list[i].charges_40[j][0])])+','+str(Destination_list[i].charges_40[j][2])[-3:]+','+'standard'+','+'40,'+str(Destination_list[i].charges_40[j][1])+','+str(Destination_list[i].charges_40[j][2])[0:-3]+'\n')

		for j in range(0,len(Destination_list[i].charges_40HC)):
			if Destination_list[i].ddf!='':
				factor=conversion[str(Destination_list[i].charges_40HC[j][2])[-3:]+'/USD']/conversion[Destination_list[i].ddf+'/USD']	
				price=float(str(Destination_list[i].charges_40HC[j][2])[0:-3])*factor
				price=round(price,2)
				output.write(str(description[str(Destination_list[i].charges_40HC[j][0])])+','+str(Destination_list[i].ddf)+','+'standard'+','+'40HC,'+str(Destination_list[i].charges_40HC[j][1])+','+str(price)+'\n')
			else:
				output.write(str(description[str(Destination_list[i].charges_40HC[j][0])])+','+str(Destination_list[i].charges_40HC[j][2])[-3:]+','+'standard'+','+'40HC,'+str(Destination_list[i].charges_40HC[j][1])+','+str(Destination_list[i].charges_40HC[j][2])[0:-3]+'\n')

		for j in range(0,len(Destination_list[i].charges_45HC)):
			if Destination_list[i].ddf!='':
				factor=conversion[str(Destination_list[i].charges_45HC[j][2])[-3:]+'/USD']/conversion[Destination_list[i].ddf+'/USD']
				price=float(str(Destination_list[i].charges_45HC[j][2])[0:-3])*factor	
				price=round(price,2)
				output.write(str(description[str(Destination_list[i].charges_45HC[j][0])])+','+str(Destination_list[i].ddf)+','+'standard'+','+'45HC,'+str(Destination_list[i].charges_45HC[j][1])+','+str(price)+'\n')
			else:
				output.write(str(description[str(Destination_list[i].charges_45HC[j][0])])+','+str(Destination_list[i].charges_45HC[j][2])[-3:]+','+'standard'+','+'45HC,'+str(Destination_list[i].charges_45HC[j][1])+','+str(Destination_list[i].charges_45HC[j][2])[0:-3]+'\n')

	index=0
	i=0
	Ocean_list=[]
	Ocean_map={}
	while i<(len(Total))-1:
		for j in range(i,len(Total)):
			if Total[j][1]!=Total[i][1] or Total[j][0]!=Total[i][0] or j==len(Total)-1:
				m=j
				break
		charges_20=[]
		charges_40=[]
		charges_40HC=[]
		charges_45HC=[]
		for k in range(i,m):
			if Total[k][9]!='' and mapping[Total[k][7]]=='Freight':	
				charges_20.append([Total[k][7],Total[k][8],Total[k][9]])
			if Total[k][10]!='' and  mapping[Total[k][7]]=='Freight':	
				charges_40.append([Total[k][7],Total[k][8],Total[k][10]])
			if Total[k][11]!='' and mapping[Total[k][7]]=='Freight':	
				charges_40HC.append([Total[k][7],Total[k][8],Total[k][11]])
			if Total[k][12]!='' and  mapping[Total[k][7]]=='Freight':	
				charges_45HC.append([Total[k][7],Total[k][8],Total[k][12]])

		if str(Total[i][0])+str(Total[i][1])  not in Ocean_map:
			Ocean_map[str(Total[i][0])+str(Total[i][1])]=index
			Ocean_list.append(Oceancharges(Total[i][0],Total[i][1],charges_20,charges_40,charges_40HC,charges_45HC))
			if len(charges_20)!=0:
				Ocean_list[index].filled_20=True
			if len(charges_40)!=0:
				Ocean_list[index].filled_40=True
			if len(charges_40HC)!=0:
				Ocean_list[index].filled_40HC=True
			if len(charges_45HC)!=0:
				Ocean_list[index].filled_45HC=True
			index=index+1	
		else:
			l=Ocean_map[str(Total[i][0])+str(Total[i][1])]
			if Ocean_list[l].filled_20==False and len(charges_20)!=0:
				Ocean_list[l].addcharges_20(charges_20)
				Ocean_list[l].filled_20=True
			if Ocean_list[l].filled_40==False and len(charges_40)!=0:
				Ocean_list[l].addcharges_40(charges_40)
				Ocean_list[l].filled_40==True
			if Ocean_list[l].filled_40HC==False and len(charges_40HC)!=0:
				Ocean_list[l].addcharges_40HC(charges_40HC)
				Ocean_list[l].filled_40HC=True
			if Ocean_list[l].filled_45HC==False and len(charges_45HC)!=0:
				Ocean_list[l].addcharges_45HC(charges_45HC)
				Ocean_list[l].filled_45HC=True

		i=m

	print('Creating Ocean Freight Charges file')

	output = open('/var/www/html/Info/Maersk/OceanFreight.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container-type'+','+'Container size'+','+'Unit'+','+'Rate/unit'+'\n')
	for i in range(0,len(Ocean_list)):
		if Ocean_list[i].filled_20:
			if Ocean_list[i].fromport in Portcodes:
				output.write('\n'+'port'+','+str(Portcodes[Ocean_list[i].fromport])+'\n')		
			else:
				output.write('\n'+'port'+','+str(Ocean_list[i].fromport).replace(","," ")+'\n')

			if Ocean_list[i].toport in Portcodes:
				output.write('port'+','+str(Portcodes[Ocean_list[i].toport])+'\n')		
			else:
				output.write('port'+','+str(Ocean_list[i].toport).replace(","," ")+'\n')
			output.write('TRANSIT TIME'+','+'\n')		
			output.write('CARRIER'+','+'Maersk'+'\n')
			output.write('SERVICE MODE'+','+str(Total[i][4])+'\n')
			output.write('ROUTING'+','+'\n')		
			output.write('Remarks'+','+'\n'+'\n'+'\n')

			for j in range(0,len(Ocean_list[i].charges_20)):
				output.write(str(description[str(Ocean_list[i].charges_20[j][0])])+','+str(Ocean_list[i].charges_20[j][2])[-3:]+','+'standard'+','+'20,'+str(Ocean_list[i].charges_20[j][1])+','+str(Ocean_list[i].charges_20[j][2])[0:-3]+'\n')


		if Ocean_list[i].filled_40:
			if Ocean_list[i].fromport in Portcodes:
				output.write('\n'+'port'+','+str(Portcodes[Ocean_list[i].fromport])+'\n')		
			else:
				output.write('\n'+'port'+','+str(Ocean_list[i].fromport).replace(","," ")+'\n')

			if Ocean_list[i].toport in Portcodes:
				output.write('port'+','+str(Portcodes[Ocean_list[i].toport])+'\n')		
			else:
				output.write('port'+','+str(Ocean_list[i].toport).replace(","," ")+'\n')
			output.write('TRANSIT TIME'+','+'\n')		
			output.write('CARRIER'+','+'Maersk'+'\n')
			output.write('SERVICE MODE'+','+str(Total[i][4])+'\n')
			output.write('ROUTING'+','+'\n')		
			output.write('Remarks'+','+'\n'+'\n'+'\n')

			for j in range(0,len(Ocean_list[i].charges_40)):
				output.write(str(description[str(Ocean_list[i].charges_40[j][0])])+','+str(Ocean_list[i].charges_40[j][2])[-3:]+','+'standard'+','+'40,'+str(Ocean_list[i].charges_40[j][1])+','+str(Ocean_list[i].charges_40[j][2])[0:-3]+'\n')		

		if Ocean_list[i].filled_40HC:
			if Ocean_list[i].fromport in Portcodes:
				output.write('\n'+'port'+','+str(Portcodes[Ocean_list[i].fromport])+'\n')		
			else:
				output.write('\n'+'port'+','+str(Ocean_list[i].fromport).replace(","," ")+'\n')

			if Ocean_list[i].toport in Portcodes:
				output.write('port'+','+str(Portcodes[Ocean_list[i].toport])+'\n')		
			else:
				output.write('port'+','+str(Ocean_list[i].toport).replace(","," ")+'\n')
			output.write('TRANSIT TIME'+','+'\n')		
			output.write('CARRIER'+','+'Maersk'+'\n')
			output.write('SERVICE MODE'+','+str(Total[i][4])+'\n')
			output.write('ROUTING'+','+'\n')		
			output.write('Remarks'+','+'\n'+'\n'+'\n')

			for j in range(0,len(Ocean_list[i].charges_40HC)):
				output.write(str(description[str(Ocean_list[i].charges_40HC[j][0])])+','+str(Ocean_list[i].charges_40HC[j][2])[-3:]+','+'standard'+','+'40HC,'+str(Ocean_list[i].charges_40HC[j][1])+','+str(Ocean_list[i].charges_40HC[j][2])[0:-3]+'\n')	


		if Ocean_list[i].filled_45HC:
			if Ocean_list[i].fromport in Portcodes:
				output.write('\n'+'port'+','+str(Portcodes[Ocean_list[i].fromport])+'\n')		
			else:
				output.write('\n'+'port'+','+str(Ocean_list[i].fromport).replace(","," ")+'\n')

			if Ocean_list[i].toport in Portcodes:
				output.write('port'+','+str(Portcodes[Ocean_list[i].toport])+'\n')		
			else:
				output.write('port'+','+str(Ocean_list[i].toport).replace(","," ")+'\n')
			output.write('TRANSIT TIME'+','+'\n')		
			output.write('CARRIER'+','+'Maersk'+'\n')
			output.write('SERVICE MODE'+','+str(Total[i][4])+'\n')
			output.write('ROUTING'+','+'\n')		
			output.write('Remarks'+','+'\n'+'\n'+'\n')

			for j in range(0,len(Ocean_list[i].charges_45HC)):
				output.write(str(description[str(Ocean_list[i].charges_45HC[j][0])])+','+str(Ocean_list[i].charges_45HC[j][2])[-3:]+','+'standard'+','+'45HC,'+str(Ocean_list[i].charges_45HC[j][1])+','+str(Ocean_list[i].charges_45HC[j][2])[0:-3]+'\n')

	fname=sys.argv[no].split("/")[-1]
	name='Outputof'+str(fname)
	workbook =Workbook('/var/www/html/outputs/Maersk/{0}.xlsx'.format(name))

	#Add an excel sheet
	worksheet1 = workbook.add_worksheet('Origin')
	x=0
	with open('/var/www/html/Info/Maersk/Origincharges.csv', newline='\n') as f:
		reader = csv.reader(f, delimiter=',')
		for row in reader:
			for y in range(len(row)):
				try:
					worksheet1.write(x,y,row[y])
				except:
					print()
			x=x+1
	#Save and create new excel file 
	#workbook.close()

	#workbook =Workbook('Output.xlsx')
	worksheet2 = workbook.add_worksheet('Destination')
	x=0
	with open('/var/www/html/Info/Maersk/Destinationcharges.csv', newline='\n') as f:
		reader = csv.reader(f, delimiter=',')
		for row in reader:
			for y in range(len(row)):
				try:
					worksheet2.write(x,y,row[y])
				except:
					print()
			x=x+1

	#Save and create new excel file 

	#workbook.close()

	#workbook =Workbook('Output.xlsx')
	worksheet3 = workbook.add_worksheet('Freight')
	x=0
	with open('/var/www/html/Info/Maersk/OceanFreight.csv', newline='\n') as f:
		reader = csv.reader(f, delimiter=',')
		for row in reader:
			for y in range(len(row)):
				try:
					worksheet3.write(x,y,row[y])
				except:
					print()
			x=x+1

	#Save and create new excel file 

	workbook.close()
