from xlrd import open_workbook
import sys
import urllib.request
import csv
from xlsxwriter import Workbook


Total=[]
for no in range(1,len(sys.argv)):
	filename = sys.argv[no]
	wb = open_workbook(filename)
	s=wb.sheet_by_index(1)
	for row in range(1,s.nrows):  # 7 from which line it starts 
		prices=[]
		for col in range(0,s.ncols):
			prices.append((s.cell(row,col).value))
		Total.append(prices)
	Origin=[]
	Destination=[]
	Freight=[]
	#print(Total)
	output = open('/var/www/html/Info/LCL/OriginChargesLCL.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Unit'+','+'Rate/unit'+','+'Minimum'+'\n'+'\n')
	for i in range(2,len(Total)):
		if Total[i][0] not in Origin:
			Origin.append(Total[i][0])
			output.write('Port /CFS'+','+Total[i][0]+'\n')
			output.write(Total[0][3]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][3])+'\n')
			output.write(Total[i][5]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][4])+'\n')
			output.write(Total[0][6]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][6])+'\n')
			output.write(Total[0][7]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][7])+'\n')
			output.write(Total[0][11]+','+Total[i][2]+','+'PER_W/M,'+str(Total[i][11])+'\n')
			output.write(Total[i][14]+','+Total[i][2]+','+'PER_W/M,'+str(Total[i][13])+','+str(Total[i][12])+'\n')
			output.write(Total[i][17]+','+Total[i][2]+','+'PER_W/M,'+str(Total[i][16])+','+str(Total[i][15])+'\n')
			output.write(Total[i][19]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][18])+'\n')
			output.write('\n')


	output = open('/var/www/html/Info/LCL/DestinationChargesLCL.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Unit'+','+'Rate/unit'+','+'Minimum'+'\n')
	for i in range(2,len(Total)):
		if Total[i][1] not in Destination:
			Destination.append(Total[i][1])
			output.write('Port /CFS'+','+Total[i][1]+'\n')
			output.write(Total[0][37]+','+Total[i][36]+','+'PER_SHIPMENT,'+str(Total[i][37])+'\n')
			output.write(Total[i][40]+','+Total[i][36]+','+'PER_SHIPMENT,'+str(Total[i][38]+Total[i][39])+'\n')
			output.write(Total[0][41]+','+Total[i][36]+','+'PER_SHIPMENT,'+str(Total[i][41])+'\n')
			output.write(Total[0][42]+','+Total[i][36]+','+'PER_W/M,'+str(Total[i][42])+','+str(Total[i][43])+'\n')
			output.write(Total[0][44]+','+Total[i][36]+','+'PER_TON,'+str(Total[i][45])+','+str(Total[i][44])+'\n')
			output.write(Total[i][49]+','+Total[i][36]+','+'PER_W/M,'+str(Total[i][48])+','+str(Total[i][47])+'\n')
			output.write(Total[i][52]+','+Total[i][36]+','+'PER_W/M,'+str(Total[i][51])+','+str(Total[i][50])+'\n')
			output.write(Total[i][54]+','+Total[i][36]+','+'PER_W/M,'+str(Total[i][53])+'\n')
			output.write('\n')
		
	output = open('/var/www/html/Info/LCL/FreightChargesLCL.csv', 'w')
	output.write('PORT FROM'+','+'PORT TO'+','+'Currency'+','+'Ocean freight charges per W/M'+'\n')
	for i in range(2,len(Total)):
		if [Total[i][0],Total[i][1]] not in Freight:
			Freight.append([Total[i][0],Total[i][1]])
			ocean=0
			for j in {21,22,23,24,25,26,28,29,31}:
				ocean=ocean+float(Total[i][j])
			output.write(Total[i][0]+','+Total[i][1]+','+Total[i][20]+','+str(ocean)+'\n')
				

				
	fname=sys.argv[no].split("/")[-1]
	name='Outputof'+str(fname)
	workbook =Workbook('/var/www/html/outputs/LCL/{0}.xlsx'.format(name))
	#to write excel(.xls) files
	#Create a new workbook object
	#Add an excel sheet
	worksheet1 = workbook.add_worksheet('Origin')
	x=0
	with open('/var/www/html/Info/LCL/OriginChargesLCL.csv', newline='\n') as f:
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

	#workbook =Workbook('OutputFCL.xlsx')
	worksheet2 = workbook.add_worksheet('Destination')
	x=0
	with open('/var/www/html/Info/LCL/DestinationChargesLCL.csv', newline='\n') as f:
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

	#workbook =Workbook('OutputFCL.xlsx')
	worksheet3 = workbook.add_worksheet('Freight')
	x=0
	with open('/var/www/html/Info/LCL/FreightChargesLCL.csv', newline='\n') as f:
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



	
	
