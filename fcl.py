from xlrd import open_workbook
import sys
import urllib.request
import csv
from xlsxwriter import Workbook
ctpye=[]
ctpye.append('20')
ctpye.append('40')
ctpye.append('40HC')
ctpye.append('45HC')
#print(ctpye)
Total=[]
for no in range(1,len(sys.argv)):
	filename = sys.argv[no]
	wb = open_workbook(filename)
	s=wb.sheet_by_index(0)
	for row in range(1,s.nrows):  # 7 from which line it starts 
		prices=[]
		for col in range(0,s.ncols):
			prices.append((s.cell(row,col).value))
		Total.append(prices)

	#print(Total)

	output = open('/var/www/html/Info/FCL/OriginchargesFCL.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container_type'+','+'Container_size'+','+'Unit'+','+'Rate/unit'+'\n'+'\n')

	for i in range(2,len(Total)):
		for j in range(0,4):
			output.write('PORT'+','+Total[i][0]+'\n')
			if Total[i][3]!=0:
				output.write(Total[0][3]+','+Total[i][2]+','+'standard,'+str(Total[1][3])+','+'PER_SHIPMENT_PER_BL,'+str(Total[i][3])+'\n')
			output.write(Total[0][3]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][4+j])+'\n')
			output.write(Total[i][9]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][8])+'\n')
			output.write(Total[0][10]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][10])+'\n')
			output.write(Total[0][11]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_BL,'+str(Total[i][11])+'\n')
			output.write(Total[0][12]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_KM_PER_CONTAINER,'+str(Total[i][12+j])+'\n')
			output.write(Total[0][16]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][16+j])+'\n')
			output.write(Total[0][20]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][20])+'\n')
			output.write(Total[i][25]+','+Total[i][2]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][21+j])+'\n')
			output.write('\n')

	output = open('/var/www/html/Info/FCL/DestinationChargesFCL.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container_type'+','+'Container_size'+','+'Unit'+','+'Rate/unit'+'\n'+'\n')

	for i in range(2,len(Total)):
		for j in range(0,4):
			output.write('PORT'+','+Total[i][1]+'\n')
			output.write(Total[0][70]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][70+j])+'\n')
			output.write(Total[i][75]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][74])+'\n')
			output.write(Total[0][76]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_SHIPMENT,'+str(Total[i][76])+'\n')
			output.write(Total[0][77]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_KM_PER_CONTAINER,'+str(Total[i][77+j])+'\n')
			output.write(Total[0][81]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_KM_PER_CONTAINER,'
			+str(Total[i][81+j]+Total[i][81+j]*Total[i][85]/100)+'\n')
			output.write(Total[i][90]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'
			+str(Total[i][86+j])+'\n')	
			output.write(Total[i][95]+','+Total[i][69]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][91+j])+'\n')
			output.write('\n')

	output = open('/var/www/html/Info/FCL/FreightChargesFCL.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Container_type'+','+'Container_size'+','+'Unit'+','+'Rate/unit'+'\n'+'\n')

	for i in range(2,len(Total)):
		for j in range(0,4):
			output.write('PORT FROM'+','+str(Total[i][0])+'\n')
			output.write('PORT TO'+','+str(Total[i][1])+'\n')
			output.write('TRANSIT TIME'+','+str(Total[i][65])+'\n')
			output.write('CARRIER'+','+str(Total[i][67])+'\n')
			output.write('SERVICE MODE'+','+Total[i][1]+'\n')
			output.write('ROUTING'+','+''+'\n')
			output.write('REMARKS'+','+''+'\n'+'\n'+'\n')
			output.write(Total[0][27]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][27+j])+'\n')
			output.write(Total[0][31]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][31+j])+'\n')
			output.write(Total[0][35]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][35+j])+'\n')
			output.write(Total[0][39]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][39+j])+'\n')
			output.write(Total[0][43]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][43+j])+'\n')
			output.write(Total[0][47]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][47+j])+'\n')
			output.write(Total[0][51]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][51+j])+'\n')
			output.write(Total[i][59]+','+Total[i][26]+','+'standard,'+str(ctpye[j])+','+'PER_CONTAINER,'+str(Total[i][55+j])+'\n')
			output.write('\n')





	#to write excel(.xls) files
	#Create a new workbook object
	fname=sys.argv[no].split("/")[-1]
	name='Outputof'+str(fname)
	workbook =Workbook('/var/www/html/outputs/FCL/{0}.xlsx'.format(name))
	#Add an excel sheet
	worksheet1 = workbook.add_worksheet('Origin')
	x=0
	with open('/var/www/html/Info/FCL/OriginchargesFCL.csv', newline='\n') as f:
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
	with open('/var/www/html/Info/FCL/DestinationChargesFCL.csv', newline='\n') as f:
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
	with open('/var/www/html/Info/FCL/FreightChargesFCL.csv', newline='\n') as f:
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

