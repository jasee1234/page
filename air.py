from xlrd import open_workbook
import sys
import urllib.request
import csv
from xlsxwriter import Workbook

Total=[]
for no in range(1,len(sys.argv)):
	filename = sys.argv[no]
	wb = open_workbook(filename)
	s=wb.sheet_by_index(2)
	for row in range(1,s.nrows):  # 7 from which line it starts 
		prices=[]
		for col in range(0,s.ncols):
			prices.append((s.cell(row,col).value))
		Total.append(prices)

	#print(Total)

	Origin=[]
	Destination=[]
	Freight=[]

	output = open('/var/www/html/Info/AIR/OriginChargesAIR.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Unit'+','+'Rate/unit'+','+'Minimum'+'\n'+'\n')


	for i in range(2,len(Total)):
		if Total[i][0] not in Origin:
			Origin.append(Total[i][0])
			output.write('AIRPORT'+','+Total[i][0]+'\n')
			output.write(Total[0][3]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][3])+'\n')
			output.write(Total[i][5]+','+Total[i][2]+','+'PER_SHIPMENT,'+str(Total[i][4])+'\n')
			output.write(Total[0][6]+','+Total[i][2]+','+'PER_KG,'+str(Total[i][7])+','+str(Total[i][6])+'\n')
			output.write(Total[0][8]+','+Total[i][2]+','+'PER_KG,'+str(Total[i][9])+','+str(Total[i][8])+'\n')
			output.write(Total[i][15]+','+Total[i][2]+','+'PER_KG,'+str(Total[i][14])+','+str(Total[i][13])+'\n')
			output.write(Total[i][18]+','+Total[i][2]+','+'PER_CHARG_KG,'+str(Total[i][17])+','+str(Total[i][16])+'\n')
			output.write('\n')

	output = open('/var/www/html/Info/AIR/FreightChargesAIR.csv', 'w')
	output.write('FROM'+','+'TO'+','+'Currency'+','+'Unit'+','+'Minimum'+','+'-45 kg'+','+'+45 kg'+',' +'+100 kg'+','+'+300 Kg'+','+'+500 kg' +','+'+1000 kg'+','+'Transit time days'+','+'Routing'+','+'carrier'+'\n'+'\n')
	for i in range(2,len(Total)):
		if [Total[i][0],Total[i][1]] not in Freight:
			Freight.append([Total[i][0],Total[i][1]])
			output.write(Total[i][0]+','+Total[i][1]+','+Total[i][19]+',PER_CHARGE_KG,'
			+str(Total[i][20])+','+str(Total[i][22])+','+str(Total[i][23])+','+str(Total[i][24])+','+str(Total[i][25])+','
			+str(Total[i][26])+','+str(Total[i][27])+','+str(Total[i][35])+','+Total[i][39]+','+Total[i][37]+'\n')


	output = open('/var/www/html/Info/AIR/DestinationChargesAIR.csv', 'w')
	output.write('Description'+','+'Currency'+','+'Unit'+','+'Rate/unit'+','+'Minimum'+'\n'+'\n')
	for i in range(2,len(Total)):
		if Total[i][1] not in Destination:
			Destination.append(Total[i][1])
			output.write('AIRPORT'+','+Total[i][1]+'\n')
			output.write(Total[i][43]+','+Total[i][40]+','+'PER_SHIPMENT,'+str(Total[i][41]+Total[i][42])+'\n')
			output.write(Total[0][44]+','+Total[i][40]+','+'PER_KG,'+str(Total[i][45])+','+str(Total[i][44])+'\n')
			output.write(Total[0][46]+','+Total[i][40]+','+'PER_KG,'+str(Total[i][47])+','+str(Total[i][46])+'\n')
			output.write(Total[i][53]+','+Total[i][40]+','+'PER_W/M,'+str(Total[i][52])+','+str(Total[i][51])+'\n')
			output.write(Total[i][56]+','+Total[i][40]+','+'PER_W/M,'+str(Total[i][55])+','+str(Total[i][54])+'\n')
			output.write('\n')

		
	fname=sys.argv[no].split("/")[-1]
	name='Outputof'+str(fname)
	workbook =Workbook('/var/www/html/outputs/AIR/{0}.xlsx'.format(name))

	#Add an excel sheet
	worksheet1 = workbook.add_worksheet('Origin')
	x=0
	with open('/var/www/html/Info/AIR/OriginChargesAIR.csv', newline='\n') as f:
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
	with open('/var/www/html/Info/AIR/DestinationChargesAIR.csv', newline='\n') as f:
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
	with open('/var/www/html/Info/AIR/FreightChargesAIR.csv', newline='\n') as f:
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

