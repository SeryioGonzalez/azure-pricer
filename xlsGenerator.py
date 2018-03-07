#!/usr/bin/python3

import datetime
import xlsxwriter

import priceReaderCompute
import priceReaderPremiumDisk
import priceReaderSiteRecovery
import priceReaderStandardDisk

workbookNamePattern = '/mnt/c/Users/segonza/Desktop/Azure-Quotes-{}.xlsx'
today = datetime.date.today().strftime('%d%m%y')

rowsForVMInput=250
firstColumnCustomerInput=3
numOfCustomerInputParams=12
conversionRate=0.82
region = 'europe-west'

firstColumnCalculations=firstColumnCustomerInput + numOfCustomerInputParams

#GET RESOURCE PRICES
computePriceMatrix = priceReaderCompute.getPriceMatrix(region)
siteRecoveryPriceMatrix = priceReaderSiteRecovery.getPriceMatrix(region)
premiumDiskPriceMatrix = priceReaderPremiumDisk.getPriceMatrix(region)
standardDiskPriceMatrix = priceReaderStandardDisk.getPriceMatrix(region)
numVmSizes = len(computePriceMatrix)

#CREATE WORKBOOKS
workbookName = workbookNamePattern.format(today)
workbookFile = workbookName
workbook = xlsxwriter.Workbook(workbookFile)
#ADD TABS
customerVMDataExcelTab = workbook.add_worksheet('customer-vm-list')
azureVMDataBaseExcelTab = workbook.add_worksheet('azure-vm-prices-base')
azureVMData1YExcelTab = workbook.add_worksheet('azure-vm-prices-1Y')
azureVMData3YExcelTab = workbook.add_worksheet('azure-vm-prices-3Y')
azureASRExcelTab = workbook.add_worksheet('azure-asr-prices')
azurePremiumDiskExcelTab  = workbook.add_worksheet('azure-premium-disk-prices')
azureStandardDiskExcelTab = workbook.add_worksheet('azure-standard-disk-prices')

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AR', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ' ]
diskSizes = ['S4', 'S6', 'S10', 'S15', 'S20', 'S30', 'S40', 'S50', 'P4', 'P6', 'P10', 'P15', 'P20', 'P30', 'P40', 'P50']

#FORMULAS
formulaTotalComputeCost="=B5*SUM(Y1:Y"+str(rowsForVMInput+1)+")"
formulaTotalDiskCost="=B5*12*(SUM(AS1:AS16) + SUM(AT1:AT"+str(rowsForVMInput+1)+")*'azure-standard-disk-prices'!C2 + SUM(AU1:AU"+str(rowsForVMInput+1)+")*'azure-premium-disk-prices'!C2)"
formulaTotalASRCost ="=B5*12*SUM(AP1:AP"+str(rowsForVMInput+1)+")*'azure-asr-prices'!A2"
formulaTotalCost="=SUM(B9:B11)"

#CHECKING ASR
formulaCheckASRPattern="=IF(AND(L{0}=\"YES\", O{0}=\"YES\"),1,\"\")"

#CHECKING STANDARD OS DISK
formulaDiskOSStandardPattern="=IF(AND(I{0}=\"STANDARD\", O{0}=\"YES\"), IF(L{0}=\"YES\",2,1) ,\"\")"

#CHECKING PREMIUM OS DISK
formulaDiskOSPremiumPattern="=IF( AND(I{0}=\"PREMIUM\",  O{0}=\"YES\"), IF(L{0}=\"YES\",2,1) ,\"\")"

#COUNTING DISKS
formulaCountS4Disk ="=SUM(Z1:Z"+str(rowsForVMInput+1)+")"
formulaCountS6Disk ="=SUM(AA1:AA"+str(rowsForVMInput+1)+")"
formulaCountS10Disk="=SUM(AB1:AB"+str(rowsForVMInput+1)+")"
formulaCountS15Disk="=SUM(AC1:AC"+str(rowsForVMInput+1)+")"
formulaCountS20Disk="=SUM(AD1:AD"+str(rowsForVMInput+1)+")"
formulaCountS30Disk="=SUM(AE1:AE"+str(rowsForVMInput+1)+")"
formulaCountS40Disk="=SUM(AF1:AF"+str(rowsForVMInput+1)+")"
formulaCountS50Disk="=SUM(AG1:AG"+str(rowsForVMInput+1)+")"
formulaCountP4Disk ="=SUM(AH1:AH"+str(rowsForVMInput+1)+")"
formulaCountP6Disk ="=SUM(AI1:AI"+str(rowsForVMInput+1)+")"
formulaCountP10Disk="=SUM(AJ1:AJ"+str(rowsForVMInput+1)+")"
formulaCountP15Disk="=SUM(AK1:AK"+str(rowsForVMInput+1)+")"
formulaCountP20Disk="=SUM(AL1:AL"+str(rowsForVMInput+1)+")"
formulaCountP30Disk="=SUM(AM1:AM"+str(rowsForVMInput+1)+")"
formulaCountP40Disk="=SUM(AN1:AO"+str(rowsForVMInput+1)+")"
formulaCountP50Disk="=SUM(AO1:AO"+str(rowsForVMInput+1)+")"
#PRICING DISKS
formulaPriceS4Disk ="=AR1*'azure-standard-disk-prices'!C2"
formulaPriceS6Disk ="=AR2*'azure-standard-disk-prices'!C3"
formulaPriceS10Disk="=AR3*'azure-standard-disk-prices'!C4"
formulaPriceS15Disk="=AR4*'azure-standard-disk-prices'!C5"
formulaPriceS20Disk="=AR5*'azure-standard-disk-prices'!C6"
formulaPriceS30Disk="=AR6*'azure-standard-disk-prices'!C7"
formulaPriceS40Disk="=AR7*'azure-standard-disk-prices'!C8"
formulaPriceS50Disk="=AR8*'azure-standard-disk-prices'!C9"
formulaPriceP4Disk ="=AR9*'azure-premium-disk-prices'!C2"
formulaPriceP6Disk ="=AR10*'azure-premium-disk-prices'!C3"
formulaPriceP10Disk="=AR11*'azure-premium-disk-prices'!C4"
formulaPriceP15Disk="=AR12*'azure-premium-disk-prices'!C5"
formulaPriceP20Disk="=AR13*'azure-premium-disk-prices'!C6"
formulaPriceP30Disk="=AR14*'azure-premium-disk-prices'!C7"
formulaPriceP40Disk="=AR15*'azure-premium-disk-prices'!C8"
formulaPriceP50Disk="=AR16*'azure-premium-disk-prices'!C9"

#CHECK ALL INPUT
formulaCheckAllInputsPattern="=IF(AND(D{0}<>\"\", E{0}<>\"\", F{0}<>\"\", G{0}<>\"\", H{0}<>\"\", I{0}<>\"\", J{0}<>\"\", K{0}<>\"\", L{0}<>\"\", M{0}<>\"\", N{0}<>\"\"),\"YES\",\"NO\")"

#VIRTUAL MACHINE FORMULAS
formulaVMBaseNamePattern    ="=IF(O{1}=\"YES\",VLOOKUP("+alphabet[firstColumnCalculations + 1]+"{1} & J{1} & K{1},'azure-vm-prices-base'!G$2:H${0}, 2, 0),\"\")"
formulaVMBaseMinPricePattern="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${0},  'azure-vm-prices-base'!A$2:A${0},\">=\"&E{1}*(100-B3)/100, 'azure-vm-prices-base'!B$2:B${0},\">=\"&F{1}*(100-B3)/100, 'azure-vm-prices-base'!D$2:D${0},J{1}, 'azure-vm-prices-base'!E$2:E${0},K{1}), _xlfn.MINIFS('azure-vm-prices-base'!I$2:I${0}, 'azure-vm-prices-base'!A$2:A${0},\">=\"&E{1}*(100-B3)/100, 'azure-vm-prices-base'!B$2:B${0},\">=\"&F{1}*(100-B3)/100, 'azure-vm-prices-base'!D$2:D${0},J{1}, 'azure-vm-prices-base'!E$2:E${0},K{1})),\"\")"

formulaVM1YNamePattern      ="=IF(O{1}=\"YES\",VLOOKUP("+alphabet[firstColumnCalculations + 3]+"{1} & J{1} & K{1},'azure-vm-prices-1Y'!G$2:H${0}, 2, 0),\"\")"
formulaVM1YMinPricePattern  ="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${0},  'azure-vm-prices-1Y'!A$2:A${0},\">=\"&E{1}*(100-B3)/100,     'azure-vm-prices-1Y'!B$2:B${0},\">=\"&F{1}*(100-B3)/100,   'azure-vm-prices-1Y'!D$2:D${0},J{1},   'azure-vm-prices-1Y'!E$2:E${0},K{1}),   _xlfn.MINIFS('azure-vm-prices-1Y'!I$2:I${0},   'azure-vm-prices-1Y'!A$2:A${0},\">=\"&E{1}*(100-B3)/100,   'azure-vm-prices-1Y'!B$2:B${0},\">=\"&F{1}*(100-B3)/100,   'azure-vm-prices-1Y'!D$2:D${0},J{1},   'azure-vm-prices-1Y'!E$2:E${0},K{1})),\"\")"
formulaVM3YNamePattern      ="=IF(O{1}=\"YES\",VLOOKUP("+alphabet[firstColumnCalculations + 5]+"{1} & J{1} & K{1},'azure-vm-prices-3Y'!G$2:H${0}, 2, 0),\"\")"
formulaVM3YMinPricePattern  ="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${0},  'azure-vm-prices-3Y'!A$2:A${0},\">=\"&E{1}*(100-B3)/100,     'azure-vm-prices-3Y'!B$2:B${0},\">=\"&F{1}*(100-B3)/100,   'azure-vm-prices-3Y'!D$2:D${0},J{1},   'azure-vm-prices-3Y'!E$2:E${0},K{1}),   _xlfn.MINIFS('azure-vm-prices-3Y'!I$2:I${0},   'azure-vm-prices-3Y'!A$2:A${0},\">=\"&E{1}*(100-B3)/100,   'azure-vm-prices-3Y'!B$2:B${0},\">=\"&F{1}*(100-B3)/100,   'azure-vm-prices-3Y'!D$2:D${0},J{1},   'azure-vm-prices-3Y'!E$2:E${0},K{1})),\"\")"
formulaVMYearPAYGPattern="=IF(O{0}=\"YES\",M{0}*Q{0}*12,\"\")"
formulaVMYear1YRIPattern="=IF(O{0}=\"YES\",S{0}*8760,\"\")"
formulaVMYear3YRIPattern="=IF(O{0}=\"YES\",U{0}*8760,\"\")"
formulaBestPricePattern ="=IF(O{0}=\"YES\",IF($B$4=\"YES\", MIN(V{0}:X{0}), V{0}),\"\")"

#STANDARD DISK FORMULAS
formulaDiskS4Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}<'azure-standard-disk-prices'!B2),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS6Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B2,G{0}<'azure-standard-disk-prices'!B3),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS10Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B3,G{0}<'azure-standard-disk-prices'!B4),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS15Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B4,G{0}<'azure-standard-disk-prices'!B5),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS20Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B5,G{0}<'azure-standard-disk-prices'!B6),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS30Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B6,G{0}<'azure-standard-disk-prices'!B7),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS40Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B7,G{0}<'azure-standard-disk-prices'!B8),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS50Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B8,G{0}<'azure-standard-disk-prices'!B9),1+IF(L{0}=\"YES\",1),\"\")"
#PREMIUM DISK FORMULAS
formulaDiskP4Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}<'azure-premium-disk-prices'!B2),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP6Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B2,G{0}<'azure-premium-disk-prices'!B3),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP10Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B3,G{0}<'azure-premium-disk-prices'!B4),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP15Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B4,G{0}<'azure-premium-disk-prices'!B5),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP20Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B5,G{0}<'azure-premium-disk-prices'!B6),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP30Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B6,G{0}<'azure-premium-disk-prices'!B7),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP40Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B7,G{0}<'azure-premium-disk-prices'!B8),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP50Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B8,G{0}<'azure-premium-disk-prices'!B9),1+IF(L{0}=\"YES\",1),\"\")"

#DEFINE FORMATS
inputHeader = workbook.add_format()
inputHeader.set_bold()
inputHeader.set_align('center')
inputHeader.set_border(1)
inputHeader.set_bg_color('#366092')

inputBody = workbook.add_format()
inputBody.set_align('center')
inputBody.set_border(1)
inputBody.set_bg_color('#b8cce4')

selecHeader = workbook.add_format()
selecHeader.set_bold()
selecHeader.set_align('center')
selecHeader.set_border(1)
selecHeader.set_bg_color('#76933c')

dataOKFormat =  workbook.add_format()
dataOKFormat.set_bold()
dataOKFormat.set_font_color('#76933c')
dataOKFormat.set_font_size(13)

selecBody = workbook.add_format()
selecBody.set_align('center')
selecBody.set_border(1)
selecBody.set_bg_color('#d8e4bc')

customerVMDataExcelTabColumnWidths = {  0:20,
										1: 5,	
									    firstColumnCustomerInput      :20, 
										firstColumnCustomerInput +  1 : 5,
										firstColumnCustomerInput +  2 : 9,
										firstColumnCustomerInput +  3 :14,
										firstColumnCustomerInput +  4 :19,
										firstColumnCustomerInput +  5 :17,
										firstColumnCustomerInput +  6 : 5,
										firstColumnCustomerInput +  7 : 5,
										firstColumnCustomerInput + 	8 : 5,
										firstColumnCustomerInput +  9 :15,
										firstColumnCustomerInput + 10 :12,
										firstColumnCustomerInput + 11 :12,
										firstColumnCalculations       :14,
										firstColumnCalculations  +  1 :13,
										firstColumnCalculations  +  2 :12,
										firstColumnCalculations  +  3 :12,
										firstColumnCalculations  +  4 :12,
										firstColumnCalculations  +  5 :12,
										firstColumnCalculations  +  6 :12,
										firstColumnCalculations  +  7 :12,
										firstColumnCalculations  +  8 :12,
										firstColumnCalculations  +  9 :15,
										firstColumnCalculations  + len(diskSizes)     : 8,
										firstColumnCalculations  + len(diskSizes) + 1 : 8,
										firstColumnCalculations  + len(diskSizes) + 2 : 8,
										firstColumnCalculations  + len(diskSizes) + 3 : 8,
										firstColumnCalculations  + len(diskSizes) + 4 : 18,
										firstColumnCalculations  + len(diskSizes) + 5 : 18											
									 }
columnDiskWidth = 12
#Width for columns from ARRAY
for columnIndex in customerVMDataExcelTabColumnWidths:
	columnWidth = customerVMDataExcelTabColumnWidths[columnIndex]
	customerVMDataExcelTab.set_column(columnIndex, columnIndex, columnWidth)
#Width for disk columns
customerVMDataExcelTab.set_column(firstColumnCalculations + 10, firstColumnCalculations + 10 + len(diskSizes) - 1, columnDiskWidth) 	

#SET FORMAT FOR ALL DATA OK
customerVMDataExcelTab.conditional_format('O1:O'+str(rowsForVMInput+1), {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '\"YES\"',
                                    'format':   dataOKFormat})

#CREATE ASSUMPTIONS
customerVMDataExcelTab.merge_range('A2:B2', 'ASSUMPTIONS', selecHeader)
	#ASSUMPTION - TOLERANCE
customerVMDataExcelTab.write(2, 0, 'PERF TOLERANCE', selecHeader)
customerVMDataExcelTab.write_number(2, 1, 0, selecBody)
customerVMDataExcelTab.data_validation('B3', {'validate': 'integer',
                                  'criteria': 'between',
                                  'minimum': 0,
                                  'maximum': 100,
                                  'input_title': 'Enter an integer:',
                                  'input_message': 'between 0 and 100 on how better % Azure perf is'})
	#ASSUMPTION - RESERVED INSTANCES
customerVMDataExcelTab.write(3, 0, 'RESERVED INSTANCES', selecHeader)
customerVMDataExcelTab.write(3, 1, 'YES', selecBody)	
customerVMDataExcelTab.data_validation('B4', {'validate': 'list','source': ['YES', 'NO']})
	#ASSUMPTION - DOLLAR TO EURO
customerVMDataExcelTab.write(4, 0, 'DOLLAR TO EURO', selecHeader)
customerVMDataExcelTab.write_number(4, 1, conversionRate, selecBody)

#CREATE COST TOTALS
customerVMDataExcelTab.merge_range('A8:B8', 'YEAR TOTALS (EUROS)', selecHeader)
customerVMDataExcelTab.write(8, 0, 'COMPUTE', selecHeader)
customerVMDataExcelTab.write(9, 0, 'STORAGE', selecHeader)
customerVMDataExcelTab.write(10,0, 'ASR',     selecHeader)
customerVMDataExcelTab.write(11, 0,'TOTAL',   selecHeader)
customerVMDataExcelTab.write_formula(8, 1, formulaTotalComputeCost, selecBody)		  
customerVMDataExcelTab.write_formula(9, 1, formulaTotalDiskCost, selecBody)
customerVMDataExcelTab.write_formula(10, 1, formulaTotalASRCost, selecBody)
customerVMDataExcelTab.write_formula(11, 1, formulaTotalCost, selecHeader)	

#CREATE DISK TOTALS
for index in range(0, len(diskSizes)):
	customerVMDataExcelTab.write(index, 42 , diskSizes[index], selecHeader)
#COUNT DISKS
customerVMDataExcelTab.write_formula(0,  43, formulaCountS4Disk, selecBody)
customerVMDataExcelTab.write_formula(1,  43, formulaCountS6Disk, selecBody)	
customerVMDataExcelTab.write_formula(2,  43, formulaCountS10Disk, selecBody)	
customerVMDataExcelTab.write_formula(3,  43, formulaCountS15Disk, selecBody)	
customerVMDataExcelTab.write_formula(4,  43, formulaCountS20Disk, selecBody)	
customerVMDataExcelTab.write_formula(5,  43, formulaCountS30Disk, selecBody)	
customerVMDataExcelTab.write_formula(6,  43, formulaCountS40Disk, selecBody)	
customerVMDataExcelTab.write_formula(7,  43, formulaCountS50Disk, selecBody)
customerVMDataExcelTab.write_formula(8,  43, formulaCountP4Disk, selecBody)
customerVMDataExcelTab.write_formula(9,  43, formulaCountP6Disk, selecBody)	
customerVMDataExcelTab.write_formula(10, 43, formulaCountP10Disk, selecBody)	
customerVMDataExcelTab.write_formula(11, 43, formulaCountP15Disk, selecBody)	
customerVMDataExcelTab.write_formula(12, 43, formulaCountP20Disk, selecBody)	
customerVMDataExcelTab.write_formula(13, 43, formulaCountP30Disk, selecBody)	
customerVMDataExcelTab.write_formula(14, 43, formulaCountP40Disk, selecBody)	
customerVMDataExcelTab.write_formula(15, 43, formulaCountP50Disk, selecBody)		 
#PRICE DISKS
customerVMDataExcelTab.write_formula(0,  44, formulaPriceS4Disk, selecBody)
customerVMDataExcelTab.write_formula(1,  44, formulaPriceS6Disk, selecBody)	
customerVMDataExcelTab.write_formula(2,  44, formulaPriceS10Disk, selecBody)	
customerVMDataExcelTab.write_formula(3,  44, formulaPriceS15Disk, selecBody)	
customerVMDataExcelTab.write_formula(4,  44, formulaPriceS20Disk, selecBody)	
customerVMDataExcelTab.write_formula(5,  44, formulaPriceS30Disk, selecBody)	
customerVMDataExcelTab.write_formula(6,  44, formulaPriceS40Disk, selecBody)	
customerVMDataExcelTab.write_formula(7,  44, formulaPriceS50Disk, selecBody)
customerVMDataExcelTab.write_formula(8,  44, formulaPriceP4Disk, selecBody)
customerVMDataExcelTab.write_formula(9,  44, formulaPriceP6Disk, selecBody)	
customerVMDataExcelTab.write_formula(10, 44, formulaPriceP10Disk, selecBody)	
customerVMDataExcelTab.write_formula(11, 44, formulaPriceP15Disk, selecBody)	
customerVMDataExcelTab.write_formula(12, 44, formulaPriceP20Disk, selecBody)	
customerVMDataExcelTab.write_formula(13, 44, formulaPriceP30Disk, selecBody)	
customerVMDataExcelTab.write_formula(14, 44, formulaPriceP40Disk, selecBody)	
customerVMDataExcelTab.write_formula(15, 44, formulaPriceP50Disk, selecBody)
	
#HEADERS - DATA FROM CUSTOMER
customerVMDataExcelTab.write(0, firstColumnCustomerInput,   'VM NAME', inputHeader)
 
customerVMDataExcelTab.write(0, firstColumnCustomerInput + 1, 'CPUs', inputHeader)

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 2, 'Mem(GB)', inputHeader)

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 3, 'DATA STORAGE', inputHeader)

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 4, 'DATA STORAGE TYPE', inputHeader)
customerVMDataExcelTab.data_validation('H2:H'+str(rowsForVMInput), {'validate': 'list','source': ['STANDARD', 'PREMIUM']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 5, 'OS STORAGE TYPE', inputHeader)
customerVMDataExcelTab.data_validation('I2:I'+str(rowsForVMInput), {'validate': 'list','source': ['STANDARD', 'PREMIUM']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 6, 'SAP', inputHeader)
customerVMDataExcelTab.data_validation('J2:J'+str(rowsForVMInput), {'validate': 'list','source': ['YES', 'NO']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 7, 'GPU', inputHeader)
customerVMDataExcelTab.data_validation('K2:K'+str(rowsForVMInput), {'validate': 'list','source': ['YES', 'NO']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 8, 'ASR', inputHeader)
customerVMDataExcelTab.data_validation('L2:L'+str(rowsForVMInput), {'validate': 'list','source': ['YES', 'NO']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 9, 'HOURS/MONTH', inputHeader)
customerVMDataExcelTab.data_validation('M2:M'+str(rowsForVMInput), {'validate': 'list','source': ['730', '264']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 10, 'USE B SERIES', inputHeader)
customerVMDataExcelTab.data_validation('N2:N'+str(rowsForVMInput), {'validate': 'list','source': ['YES', 'NO']})

customerVMDataExcelTab.write(0, firstColumnCustomerInput + 11, 'ALL DATA OK', inputHeader)

#HEADERS - CALCULATIONS
customerVMDataExcelTab.write(0, firstColumnCalculations + 0, 'BEST SIZE PAYG', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 1, 'PRICE(H) PAYG', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 2, 'BEST SIZE 1Y', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 3, 'PRICE(H) 1Y', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 4, 'BEST SIZE 3Y', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 5, 'PRICE(H) 3Y', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 6, 'YEAR PAYG', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 7, '1Y RI', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 8, '3Y RI', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 9, 'BEST PRICE', selecHeader)

#PUT HEADERS FOR MANAGED DISKS
for index in range(0, len(diskSizes)):
	customerVMDataExcelTab.write(0, firstColumnCalculations + 10 + index, 'M. DISK - '+diskSizes[index], selecHeader)

customerVMDataExcelTab.write(0, firstColumnCalculations + 10 + len(diskSizes), 'HAS ASR', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 10 + len(diskSizes) + 4, 'OS DISK STANDARD', selecHeader)
customerVMDataExcelTab.write(0, firstColumnCalculations + 10 + len(diskSizes) + 5, 'OS DISK PREMIUM', selecHeader)

#CONTENT STYLE FOR EVERY ROW FOR CUSTOMER INPUT 
for rowIndex in range(1,rowsForVMInput):
	for columnIndex in range(firstColumnCustomerInput,firstColumnCustomerInput + numOfCustomerInputParams):
		customerVMDataExcelTab.write(rowIndex, columnIndex, '', inputBody)
		
#PUT DEFAULT VALUES
for rowIndex in range(1,rowsForVMInput):
#PUT ALL DATA STORAGE TYPE TO STANDARD
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  4, 'STANDARD', inputBody)	
#PUT ALL OS STORAGE TYPE TO STANDARD
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  5, 'STANDARD', inputBody)	
#PUT ALL SAP VMs TO NO
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  6, 'NO', inputBody)
#PUT ALL GPU VMs TO NO
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  7, 'NO', inputBody)
#PUT ALL ASR TO NO
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  8, 'NO', inputBody)	
#PUT ALL MONTHLY HOURS TO 730
	customerVMDataExcelTab.write_number(rowIndex, firstColumnCustomerInput +  9, 730, inputBody)
#PUT ALL BURSTABLE VMs TO NO
	customerVMDataExcelTab.write(rowIndex, firstColumnCustomerInput +  10, "NO", inputBody)
	
#FORMULAS AND STYLE FOR CALCULATIONS
for rowIndex in range(1,rowsForVMInput):		
	formulaCheckAllInputs =formulaCheckAllInputsPattern.format( rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations - 1, formulaCheckAllInputs,   inputBody)	

	formulaVMBaseName  =formulaVMBaseNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 0, formulaVMBaseName,   selecBody)		
	
	formulaVMBaseMinPrice=formulaVMBaseMinPricePattern.format(numVmSizes, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 1, formulaVMBaseMinPrice, selecBody)
	
	formulaVM1YName  =formulaVM1YNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 2, formulaVM1YName,   selecBody)		
	
	formulaVM1YMinPrice=formulaVM1YMinPricePattern.format(numVmSizes, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 3, formulaVM1YMinPrice, selecBody)

	formulaVM3YName  =formulaVM3YNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 4, formulaVM3YName,   selecBody)		
	
	formulaVM3YMinPrice=formulaVM3YMinPricePattern.format(numVmSizes, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 5, formulaVM3YMinPrice, selecBody)

	formulaVMYearPAYG=formulaVMYearPAYGPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 6, formulaVMYearPAYG, selecBody)

	formulaVMYear1YRI=formulaVMYear1YRIPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 7, formulaVMYear1YRI, selecBody)

	formulaVMYear3YRI=formulaVMYear3YRIPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 8, formulaVMYear3YRI, selecBody)
	
	formulaBestPrice=formulaBestPricePattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 9, formulaBestPrice, selecBody)	
#STANDARD DISKS	
	formulaDiskS4=formulaDiskS4Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 10, formulaDiskS4, selecBody)	

	formulaDiskS6=formulaDiskS6Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 11, formulaDiskS6, selecBody)		
	
	formulaDiskS10=formulaDiskS10Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 12, formulaDiskS10, selecBody)			
	
	formulaDiskS15=formulaDiskS15Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 13, formulaDiskS15, selecBody)		
	
	formulaDiskS20=formulaDiskS20Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 14, formulaDiskS20, selecBody)		
	
	formulaDiskS30=formulaDiskS30Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 15, formulaDiskS30, selecBody)		
	
	formulaDiskS40=formulaDiskS40Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 16, formulaDiskS40, selecBody)		
	
	formulaDiskS50=formulaDiskS50Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 17, formulaDiskS50, selecBody)
#MANAGED DISKS	
	formulaDiskP4=formulaDiskP4Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 18, formulaDiskP4, selecBody)	

	formulaDiskP6=formulaDiskP6Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 19, formulaDiskP6, selecBody)		
	
	formulaDiskP10=formulaDiskP10Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 20, formulaDiskP10, selecBody)			
	
	formulaDiskP15=formulaDiskP15Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 21, formulaDiskP15, selecBody)		
	
	formulaDiskP20=formulaDiskP20Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 22, formulaDiskP20, selecBody)		
	
	formulaDiskP30=formulaDiskP30Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 23, formulaDiskP30, selecBody)		
	
	formulaDiskP40=formulaDiskP40Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 24, formulaDiskP40, selecBody)		
	
	formulaDiskP50=formulaDiskP50Pattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 25, formulaDiskP50, selecBody)
#ASR	
	formulaCheckASR=formulaCheckASRPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 26, formulaCheckASR, selecBody)
#COUNT OS STANDARD	
	formulaDiskOSStandard=formulaDiskOSStandardPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 30, formulaDiskOSStandard, selecBody)
#COUNT OS PREMIUM
	formulaDiskOSPremium=formulaDiskOSPremiumPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 31, formulaDiskOSPremium, selecBody)

	
############ OTHER TABS	
#PUT DATA FROM APIs
	#BASE
azureVMDataBaseExcelTab.write(0, 0, 'CPUs', inputHeader)
azureVMDataBaseExcelTab.write(0, 1, 'Mem(GB)', inputHeader)
azureVMDataBaseExcelTab.set_column(2, 2, 12) 
azureVMDataBaseExcelTab.write(0, 2, 'Price/Hour', inputHeader)
azureVMDataBaseExcelTab.set_column(3, 5, 9)
azureVMDataBaseExcelTab.write(0, 3, 'SAP', inputHeader)
azureVMDataBaseExcelTab.write(0, 4, 'GPU', inputHeader)
azureVMDataBaseExcelTab.write(0, 5, 'Burstable', inputHeader)
azureVMDataBaseExcelTab.set_column(6, 6, 0) 
azureVMDataBaseExcelTab.set_column(7, 7, 20)
azureVMDataBaseExcelTab.write(0, 7, 'VM SIZE NAME', inputHeader)
azureVMDataBaseExcelTab.set_column(8, 8, 0)

	#1Y
azureVMData1YExcelTab.write(0, 0, 'CPUs', inputHeader)
azureVMData1YExcelTab.write(0, 1, 'Mem(GB)', inputHeader)
azureVMData1YExcelTab.set_column(2, 2, 12) 
azureVMData1YExcelTab.write(0, 2, 'Price/Hour', inputHeader)
azureVMData1YExcelTab.set_column(3, 5, 9)
azureVMData1YExcelTab.write(0, 3, 'SAP', inputHeader)
azureVMData1YExcelTab.write(0, 4, 'GPU', inputHeader)
azureVMData1YExcelTab.write(0, 5, 'Burstable', inputHeader)
azureVMData1YExcelTab.set_column(6, 6, 0) 
azureVMData1YExcelTab.set_column(7, 7, 20)
azureVMData1YExcelTab.write(0, 7, 'VM SIZE NAME', inputHeader)
azureVMData1YExcelTab.set_column(8, 8, 0) 

	#3Y
azureVMData3YExcelTab.write(0, 0, 'CPUs', inputHeader)
azureVMData3YExcelTab.write(0, 1, 'Mem(GB)', inputHeader)
azureVMData3YExcelTab.set_column(2, 2, 12) 
azureVMData3YExcelTab.write(0, 2, 'Price/Hour', inputHeader)
azureVMData3YExcelTab.set_column(3, 5, 9)
azureVMData3YExcelTab.write(0, 3, 'SAP', inputHeader)
azureVMData3YExcelTab.write(0, 4, 'GPU', inputHeader)
azureVMData3YExcelTab.write(0, 5, 'Burstable', inputHeader)
azureVMData3YExcelTab.set_column(6, 6, 0) 
azureVMData3YExcelTab.set_column(7, 7, 20)
azureVMData3YExcelTab.write(0, 7, 'VM SIZE NAME', inputHeader)
azureVMData3YExcelTab.set_column(8, 8, 0) 

currentLineBase = 1
currentLine1Y = 1
currentLine3Y = 1

	#DUMP API DATA
for size in sorted(computePriceMatrix):

	cpus = computePriceMatrix[size]['cpu']
	mem  = computePriceMatrix[size]['ram']
	priceBase = computePriceMatrix[size]['payg']
	price1Y=computePriceMatrix[size]['1y']
	price3Y=computePriceMatrix[size]['3y']	
	name=size
	SAP=computePriceMatrix[size]['sap']
	GPU=computePriceMatrix[size]['gpu']
	isBurstable=computePriceMatrix[size]['burstable']
	
	#WRITE TO THE BASIC
	priceBase=float(priceBase)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 0, cpus,      inputBody)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 1, mem,       inputBody)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 2, priceBase, inputBody)
	azureVMDataBaseExcelTab.write(currentLineBase, 3, SAP, inputBody)
	azureVMDataBaseExcelTab.write(currentLineBase, 4, GPU, inputBody)
	azureVMDataBaseExcelTab.write(currentLineBase, 5, isBurstable, inputBody)
	auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLineBase+1)
	azureVMDataBaseExcelTab.write_formula(currentLineBase, 6, auxFormula, inputBody)
	azureVMDataBaseExcelTab.write(currentLineBase, 7, name, inputBody)
	
	noBurstablePriceBase = priceBase
	noBurstablePrice1Y = price1Y
	noBurstablePrice3Y = price3Y
	if isBurstable == "YES":
		noBurstablePriceBase += 10000
		noBurstablePrice1Y += 10000
		noBurstablePrice3Y += 10000
		
	azureVMDataBaseExcelTab.write(currentLineBase, 8, noBurstablePriceBase, inputBody)	
	
	currentLineBase += 1	

	if price1Y != 1000000:
		price1Y=float(price1Y)
		azureVMData1YExcelTab.write_number(currentLine1Y, 0, cpus,      inputBody)
		azureVMData1YExcelTab.write_number(currentLine1Y, 1, mem,       inputBody)
		azureVMData1YExcelTab.write_number(currentLine1Y, 2, price1Y, inputBody)
		azureVMData1YExcelTab.write(currentLine1Y, 3, SAP, inputBody)
		azureVMData1YExcelTab.write(currentLine1Y, 4, GPU, inputBody)
		azureVMData1YExcelTab.write(currentLine1Y, 5, isBurstable, inputBody)
		auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLine1Y+1)
		azureVMData1YExcelTab.write_formula(currentLine1Y, 6, auxFormula, inputBody)
		azureVMData1YExcelTab.write(currentLine1Y, 7, name, inputBody)
		azureVMData1YExcelTab.write(currentLine1Y, 8, noBurstablePrice1Y, inputBody)	
		currentLine1Y += 1		
	
	if price3Y != 1000000:
		price3Y=float(price3Y)
		azureVMData3YExcelTab.write_number(currentLine3Y, 0, cpus,      inputBody)
		azureVMData3YExcelTab.write_number(currentLine3Y, 1, mem,       inputBody)
		azureVMData3YExcelTab.write_number(currentLine3Y, 2, price3Y, inputBody)
		azureVMData3YExcelTab.write(currentLine3Y, 3, SAP, inputBody)
		azureVMData3YExcelTab.write(currentLine3Y, 4, GPU, inputBody)
		azureVMData3YExcelTab.write(currentLine3Y, 5, isBurstable, inputBody)
		auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLine3Y+1)
		azureVMData3YExcelTab.write_formula(currentLine3Y, 6, auxFormula, inputBody)
		azureVMData3YExcelTab.write(currentLine3Y, 7, name, inputBody)
		azureVMData3YExcelTab.write(currentLine3Y, 8, noBurstablePrice3Y, inputBody)	
		currentLine3Y += 1		

#ASR
azureASRExcelTab.write(0, 0, 'ASR Azure to Azure', inputHeader)
azureASRExcelTab.set_column(0, 0, 25)
azureASRExcelTab.write_number(1, 0, 25,      inputBody)
for size in siteRecoveryPriceMatrix:
	priceBase = siteRecoveryPriceMatrix[size]
	azureASRExcelTab.write_number(1, 0, priceBase,      inputBody)

#PREMIUM STORAGE
azurePremiumDiskExcelTab.write(0, 0, 'Disk Size', inputHeader)
azurePremiumDiskExcelTab.set_column(0, 0, 25)
azurePremiumDiskExcelTab.set_column(1, 2, 15)
azurePremiumDiskExcelTab.write(0, 1, 'Capacity', inputHeader)
azurePremiumDiskExcelTab.write(0, 2, 'Cost', inputHeader)

currentLine = 1
for size in  sorted(premiumDiskPriceMatrix):
	priceBase = premiumDiskPriceMatrix[size]['price']
	name = premiumDiskPriceMatrix[size]['name']
	
	azurePremiumDiskExcelTab.write(currentLine, 0, name,      inputBody)
	azurePremiumDiskExcelTab.write(currentLine, 1, size,  inputBody)
	azurePremiumDiskExcelTab.write(currentLine, 2, priceBase, inputBody)
	currentLine += 1	
		
#STANDARD STORAGE 
azureStandardDiskExcelTab.write(0, 0, 'Disk Size', inputHeader)
azureStandardDiskExcelTab.set_column(0, 0, 25)
azureStandardDiskExcelTab.set_column(1, 2, 15)
azureStandardDiskExcelTab.write(0, 1, 'Capacity', inputHeader)
azureStandardDiskExcelTab.write(0, 2, 'Cost', inputHeader)

currentLine = 1
for size in  sorted(standardDiskPriceMatrix):
	priceBase = standardDiskPriceMatrix[size]['price']
	name = standardDiskPriceMatrix[size]['name']
	
	azureStandardDiskExcelTab.write(currentLine, 0, name,      inputBody)
	azureStandardDiskExcelTab.write(currentLine, 1, size,  inputBody)
	azureStandardDiskExcelTab.write(currentLine, 2, priceBase, inputBody)
	currentLine += 1	

workbook.close()    