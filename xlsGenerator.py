#!/usr/bin/python3

import datetime
import xlsxwriter

from xlsStructure import xlsStructure as xls
import priceReaderCompute
import priceReaderManagedDisk
import priceReaderSiteRecovery

workbookNamePattern = '/mnt/c/Users/segonza/Desktop/Azure-Quotes-{}.xlsx'
today = datetime.date.today().strftime('%d%m%y')
region = 'europe-west'

numOfCustomerInputParams=12
firstColumnCalculations=xls.firstColumnCustomerInput + numOfCustomerInputParams

#KEY COLUMN INDEX
dataOKColumnLetter=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ALL DATA OK']['index']))
ASRColumnLetter=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ASR']['index']))

totalNumDisks = len(priceReaderManagedDisk.standardDiskSizes) + len(priceReaderManagedDisk.premiumDiskSizes)

#1 - GET RESOURCE PRICES
computePriceMatrix = priceReaderCompute.getPriceMatrix(region)
siteRecoveryPriceMatrix = priceReaderSiteRecovery.getPriceMatrix(region)
premiumDiskPriceMatrix  = priceReaderManagedDisk.getPriceMatrixPremium(region)
standardDiskPriceMatrix = priceReaderManagedDisk.getPriceMatrixStandard(region)
numVmSizes = len(computePriceMatrix)

#2 - CREATE WORKBOOKS
workbookFile = workbookNamePattern.format(today)
workbook = xlsxwriter.Workbook(workbookFile)

#3 - ADD TABS
customerVMDataExcelTab = workbook.add_worksheet('customer-vm-list')
azureVMDataBaseExcelTab = workbook.add_worksheet('azure-vm-prices-base')
azureVMData1YExcelTab = workbook.add_worksheet('azure-vm-prices-1Y')
azureVMData3YExcelTab = workbook.add_worksheet('azure-vm-prices-3Y')
azureASRExcelTab = workbook.add_worksheet('azure-asr-prices')
azurePremiumDiskExcelTab  = workbook.add_worksheet('azure-premium-disk-prices')
azureStandardDiskExcelTab = workbook.add_worksheet('azure-standard-disk-prices')

#4 - DEFINE FORMATS
inputHeaderStyle = workbook.add_format()
inputHeaderStyle.set_bold()
inputHeaderStyle.set_align('center')
inputHeaderStyle.set_border(1)
inputHeaderStyle.set_bg_color('#366092')
inputBodyStyle = workbook.add_format()
inputBodyStyle.set_align('center')
inputBodyStyle.set_border(1)
inputBodyStyle.set_bg_color('#b8cce4')
selectHeaderStyle = workbook.add_format()
selectHeaderStyle.set_bold()
selectHeaderStyle.set_align('center')
selectHeaderStyle.set_border(1)
selectHeaderStyle.set_bg_color('#76933c')
dataOKFormat =  workbook.add_format()
dataOKFormat.set_bold()
dataOKFormat.set_font_color('#76933c')
dataOKFormat.set_font_size(13)
selecBody = workbook.add_format()
selecBody.set_align('center')
selecBody.set_border(1)
selecBody.set_bg_color('#d8e4bc')

########################################################
#####################PUT DATA BLOCKS####################
########################################################

#1st COLUMN WIDTH
customerVMDataExcelTab.set_column(0, 0, xls.firstColumnWidth) 	

#5 - BLOCK 1 - CREATE ASSUMPTIONS
	#GET ASSUMPTION VALUE CELLS
perfGainValueCell=xls.getAssumptionValueCell('PERF')
resInstValueCell=xls.getAssumptionValueCell('RESINST')
currencyRateCell=xls.getAssumptionValueCell('USD2EURO')

	#CALCULATE AND PUT HEADER
firstColumnLetter=xls.getColumnLetterFromIndex(xls.assumptions['firstCellColumn'])
firstRowLetter=xls.assumptions['firstCellRow'] + 1
lastColumnLetter=xls.getColumnLetterFromIndex(xls.assumptions['firstCellColumn'] + xls.assumptions['header']['width'] - 1)
headerRange='{0}{1}:{2}{1}'.format(firstColumnLetter, firstRowLetter, lastColumnLetter)
customerVMDataExcelTab.merge_range(headerRange, xls.assumptions['header']['title'], selectHeaderStyle)

	#ASSUMPTION - TOLERANCE
category='PERF'
name=xls.assumptions['rows'][category]['name']
row=xls.assumptions['firstCellRow'] + xls.assumptions['rows'][category]['order']
defaultValue=xls.assumptions['rows'][category]['default']
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'], name, selectHeaderStyle)
customerVMDataExcelTab.write_number(row, xls.assumptions['firstCellColumn'] + 1, defaultValue, selecBody)
customerVMDataExcelTab.data_validation(perfGainValueCell, {'validate': 'integer', 'criteria': 'between',
                                  'minimum': 0, 'maximum': 100, 'input_title': 'Enter an integer:',
                                  'input_message': 'between 0 and 100 on how better % Azure perf is'})

	#ASSUMPTION - RESERVED INSTANCES
category='RESINST'
name=xls.assumptions['rows'][category]['name']
row=xls.assumptions['firstCellRow'] + xls.assumptions['rows'][category]['order']
defaultValue=xls.assumptions['rows'][category]['default']
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'], name, selectHeaderStyle)
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'] + 1, defaultValue, selecBody)	
customerVMDataExcelTab.data_validation(resInstValueCell, {'validate': 'list','source': ['YES', 'NO']})

	#ASSUMPTION - DOLLAR TO EURO
category='USD2EURO'
name=xls.assumptions['rows'][category]['name']
row=xls.assumptions['firstCellRow'] + xls.assumptions['rows'][category]['order']
defaultValue=xls.assumptions['rows'][category]['default']
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'], name, selectHeaderStyle)
customerVMDataExcelTab.write_number(row, xls.assumptions['firstCellColumn'] + 1, defaultValue, selecBody)	

#5 - BLOCK 2 - CREATE CUSTOMER INPUT COLUMNS
for column in xls.customerInputColumns['columns']:
	#GET COLUMN DATA
	columnWidth = xls.customerInputColumns['columns'][column]['width']
	columnName  = xls.customerInputColumns['columns'][column]['alias']
	columnPositon = xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns'][column]['index'])
	columnDefaultValue = xls.customerInputColumns['columns'][column]['default']
	
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnPositon, columnPositon, columnWidth)
	#SET HEADER
	customerVMDataExcelTab.write(xls.customerInputColumns['firstCellRow'], columnPositon, columnName, inputHeaderStyle)
	
	#SET DEFAULT VALUE OR NONE
	for rowIndex in range(xls.customerInputColumns['firstCellRow'] + 1, xls.rowsForVMInput):
		customerVMDataExcelTab.write(rowIndex, columnPositon, columnDefaultValue, inputBodyStyle)	
	
	#SET INPUT RESTRICTIONS IF ANY
	try:
		columnValidationList = xls.customerInputColumns['columns'][column]['validationList']
		customerVMDataExcelTab.data_validation(xls.customerInputColumns['firstCellRow'] + 1, columnPositon, xls.rowsForVMInput, columnPositon, {'validate': 'list','source': columnValidationList})
	except:
		pass
		
#6 - BLOCK 3 - CREATE VM CALCULATION COLUMNS
columnBestVMPrice = xls.getVMCalculationColumn('BEST PRICE')
for column in xls.VMCalculationColumns['columns']:
	#GET COLUMN DATA
	columnWidth = xls.VMCalculationColumns['columns'][column]['width']
	columnName  = xls.VMCalculationColumns['columns'][column]['alias']
	columnPositon = xls.getCalculationColumnPositionInExcel(xls.VMCalculationColumns['columns'][column]['index'])
	
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnPositon, columnPositon, columnWidth)
	#SET HEADER
	customerVMDataExcelTab.write(xls.VMCalculationColumns['firstCellRow'], columnPositon, columnName, selectHeaderStyle)
	
#7 - BLOCK 4 - CREATE DATA DISK CALCULATION COLUMNS
	#GET  DATA
dataDiskFirstColumn = xls.managedDataDiskColumns['firstColumnIndex']
dataDiskPrefix = xls.managedDataDiskColumns['prefix']
dataDiskColumnWidth = xls.managedDataDiskColumns['width']

	#STANDARD DISKS
for diskIndex in range(xls.managedDataDiskColumns['firstCellRow'], len(priceReaderManagedDisk.standardDiskSizes) ):
	columnIndex = dataDiskFirstColumn + diskIndex
	diskName=priceReaderManagedDisk.standardDiskSizes[diskIndex]
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnIndex, columnIndex, dataDiskColumnWidth) 		
	#SET HEADER
	customerVMDataExcelTab.write(xls.managedDataDiskColumns['firstCellRow'], columnIndex, dataDiskPrefix+diskName, selectHeaderStyle)
	
	#PREMIUM DISKS
dataDiskFirstColumn = columnIndex + 1
for diskIndex in range(xls.managedDataDiskColumns['firstCellRow'], len(priceReaderManagedDisk.premiumDiskSizes) ):
	columnIndex = dataDiskFirstColumn + diskIndex
	diskName=priceReaderManagedDisk.premiumDiskSizes[diskIndex]
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnIndex, columnIndex, dataDiskColumnWidth) 		
	#SET HEADER
	customerVMDataExcelTab.write(xls.managedDataDiskColumns['firstCellRow'], columnIndex, dataDiskPrefix+diskName, selectHeaderStyle)

	
#7 - BLOCK 5 - ASR CALCULATION COLUMNS
#SET WIDTH
customerVMDataExcelTab.set_column(xls.ASRColumn['firstColumnIndex'], xls.ASRColumn['firstColumnIndex'], xls.ASRColumn['width'])
#SET HEADER
customerVMDataExcelTab.write(xls.ASRColumn['firstCellRow'], xls.ASRColumn['firstColumnIndex'], xls.ASRColumn['name'] , selectHeaderStyle)
#ROWS
formulaCheckASRPattern="=IF(AND({0}{1}=\"YES\", {2}{1}=\"YES\"),1,\"\")"	
for rowIndex in range(1,xls.rowsForVMInput):		
	formulaCheckASR=formulaCheckASRPattern.format(ASRColumnLetter,rowIndex+1,dataOKColumnLetter)
	customerVMDataExcelTab.write_formula(rowIndex, xls.ASRColumn['firstColumnIndex'], formulaCheckASR, selecBody)

	
#FORMULAS

#BLOCK 9
formulaTotalComputeCost="={0}*SUM({1}1:{1}{2})".format(currencyRateCell, columnBestVMPrice, xls.rowsForVMInput + 1)
formulaTotalASRCost ="={0}*12*SUM({1}{2}:{1}{3})*'azure-asr-prices'!A2".format(currencyRateCell, xls.getColumnLetterFromIndex(xls.ASRColumn['firstColumnIndex']), xls.ASRColumn['firstCellRow'] + 1, xls.rowsForVMInput + 1)
#######################################################################
#######################################################################
#######################################################################
#######################################################################
#######################################################################

#FORMULAS COUNTING DISKS
formulaCountS4Disk ="=SUM(Z1:Z"+str(xls.rowsForVMInput+1)+")"
formulaCountS6Disk ="=SUM(AA1:AA"+str(xls.rowsForVMInput+1)+")"
formulaCountS10Disk="=SUM(AB1:AB"+str(xls.rowsForVMInput+1)+")"
formulaCountS15Disk="=SUM(AC1:AC"+str(xls.rowsForVMInput+1)+")"
formulaCountS20Disk="=SUM(AD1:AD"+str(xls.rowsForVMInput+1)+")"
formulaCountS30Disk="=SUM(AE1:AE"+str(xls.rowsForVMInput+1)+")"
formulaCountS40Disk="=SUM(AF1:AF"+str(xls.rowsForVMInput+1)+")"
formulaCountS50Disk="=SUM(AG1:AG"+str(xls.rowsForVMInput+1)+")"
formulaCountP4Disk ="=SUM(AH1:AH"+str(xls.rowsForVMInput+1)+")"
formulaCountP6Disk ="=SUM(AI1:AI"+str(xls.rowsForVMInput+1)+")"
formulaCountP10Disk="=SUM(AJ1:AJ"+str(xls.rowsForVMInput+1)+")"
formulaCountP15Disk="=SUM(AK1:AK"+str(xls.rowsForVMInput+1)+")"
formulaCountP20Disk="=SUM(AL1:AL"+str(xls.rowsForVMInput+1)+")"
formulaCountP30Disk="=SUM(AM1:AM"+str(xls.rowsForVMInput+1)+")"
formulaCountP40Disk="=SUM(AN1:AO"+str(xls.rowsForVMInput+1)+")"
formulaCountP50Disk="=SUM(AO1:AO"+str(xls.rowsForVMInput+1)+")"

#FORMULAS PRICING DISKS
formulaPriceS4Disk ="=B16*'azure-standard-disk-prices'!C2"
formulaPriceS6Disk ="=B17*'azure-standard-disk-prices'!C3"
formulaPriceS10Disk="=B18*'azure-standard-disk-prices'!C4"
formulaPriceS15Disk="=B19*'azure-standard-disk-prices'!C5"
formulaPriceS20Disk="=B20*'azure-standard-disk-prices'!C6"
formulaPriceS30Disk="=B21*'azure-standard-disk-prices'!C7"
formulaPriceS40Disk="=B22*'azure-standard-disk-prices'!C8"
formulaPriceS50Disk="=B23*'azure-standard-disk-prices'!C9"
formulaPriceP4Disk ="=B24*'azure-premium-disk-prices'!C2"
formulaPriceP6Disk ="=B25*'azure-premium-disk-prices'!C3"
formulaPriceP10Disk="=B26*'azure-premium-disk-prices'!C4"
formulaPriceP15Disk="=B27*'azure-premium-disk-prices'!C5"
formulaPriceP20Disk="=B28*'azure-premium-disk-prices'!C6"
formulaPriceP30Disk="=B29*'azure-premium-disk-prices'!C7"
formulaPriceP40Disk="=B30*'azure-premium-disk-prices'!C8"
formulaPriceP50Disk="=B31*'azure-premium-disk-prices'!C9"

#COUNT DISKS
customerVMDataExcelTab.write_formula(15, 1, formulaCountS4Disk, selecBody)
customerVMDataExcelTab.write_formula(16, 1, formulaCountS6Disk, selecBody)	
customerVMDataExcelTab.write_formula(17, 1, formulaCountS10Disk, selecBody)	
customerVMDataExcelTab.write_formula(18, 1, formulaCountS15Disk, selecBody)	
customerVMDataExcelTab.write_formula(19, 1, formulaCountS20Disk, selecBody)	
customerVMDataExcelTab.write_formula(20, 1, formulaCountS30Disk, selecBody)	
customerVMDataExcelTab.write_formula(21, 1, formulaCountS40Disk, selecBody)	
customerVMDataExcelTab.write_formula(22, 1, formulaCountS50Disk, selecBody)
customerVMDataExcelTab.write_formula(23, 1, formulaCountP4Disk, selecBody)
customerVMDataExcelTab.write_formula(24, 1, formulaCountP6Disk, selecBody)	
customerVMDataExcelTab.write_formula(25, 1, formulaCountP10Disk, selecBody)	
customerVMDataExcelTab.write_formula(26, 1, formulaCountP15Disk, selecBody)	
customerVMDataExcelTab.write_formula(27, 1, formulaCountP20Disk, selecBody)	
customerVMDataExcelTab.write_formula(28, 1, formulaCountP30Disk, selecBody)	
customerVMDataExcelTab.write_formula(29, 1, formulaCountP40Disk, selecBody)	
customerVMDataExcelTab.write_formula(30, 1, formulaCountP50Disk, selecBody)		 
#PRICE DISKS
customerVMDataExcelTab.write_formula(15, 2, formulaPriceS4Disk, selecBody)
customerVMDataExcelTab.write_formula(16, 2, formulaPriceS6Disk, selecBody)	
customerVMDataExcelTab.write_formula(17, 2, formulaPriceS10Disk, selecBody)	
customerVMDataExcelTab.write_formula(18, 2, formulaPriceS15Disk, selecBody)	
customerVMDataExcelTab.write_formula(19, 2, formulaPriceS20Disk, selecBody)	
customerVMDataExcelTab.write_formula(20, 2, formulaPriceS30Disk, selecBody)	
customerVMDataExcelTab.write_formula(21, 2, formulaPriceS40Disk, selecBody)	
customerVMDataExcelTab.write_formula(22, 2, formulaPriceS50Disk, selecBody)
customerVMDataExcelTab.write_formula(23, 2, formulaPriceP4Disk, selecBody)
customerVMDataExcelTab.write_formula(24, 2, formulaPriceP6Disk, selecBody)	
customerVMDataExcelTab.write_formula(25, 2, formulaPriceP10Disk, selecBody)	
customerVMDataExcelTab.write_formula(26, 2, formulaPriceP15Disk, selecBody)	
customerVMDataExcelTab.write_formula(27, 2, formulaPriceP20Disk, selecBody)	
customerVMDataExcelTab.write_formula(28, 2, formulaPriceP30Disk, selecBody)	
customerVMDataExcelTab.write_formula(29, 2, formulaPriceP40Disk, selecBody)	
customerVMDataExcelTab.write_formula(30, 2, formulaPriceP50Disk, selecBody)


formulaTotalDiskCost="={0}*12*( SUM(C16:C31) + SUM(AQ1:AQ{1})*'azure-standard-disk-prices'!C2 + SUM(AR1:AR{1})*'azure-premium-disk-prices'!C2)".format(currencyRateCell, xls.rowsForVMInput + 1)
formulaTotalCost="=SUM(B9:B11)"


#CHECKING STANDARD OS DISK
formulaDiskOSStandardPattern="=IF(AND(I{0}=\"STANDARD\", O{0}=\"YES\"), IF(L{0}=\"YES\",2,1) ,\"\")"

#CHECKING PREMIUM OS DISK
formulaDiskOSPremiumPattern="=IF( AND(I{0}=\"PREMIUM\",  O{0}=\"YES\"), IF(L{0}=\"YES\",2,1) ,\"\")"




#CHECK ALL INPUT
formulaCheckAllInputsPattern="=IF(AND(D{0}<>\"\", E{0}<>\"\", F{0}<>\"\", G{0}<>\"\", H{0}<>\"\", I{0}<>\"\", J{0}<>\"\", K{0}<>\"\", L{0}<>\"\", M{0}<>\"\", N{0}<>\"\"),\"YES\",\"NO\")"

#VIRTUAL MACHINE FORMULAS
formulaVMBaseNamePattern    ="=IF(O{1}=\"YES\",VLOOKUP("+xls.alphabet[firstColumnCalculations + 1]+"{1} & J{1} & K{1},'azure-vm-prices-base'!G$2:H${0}, 2, 0),\"\")"
formulaVMBaseMinPricePattern="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${0},  'azure-vm-prices-base'!A$2:A${0},\">=\"&E{1}*(100-{2})/100, 'azure-vm-prices-base'!B$2:B${0},\">=\"&F{1}*(100-{2})/100, 'azure-vm-prices-base'!D$2:D${0},J{1}, 'azure-vm-prices-base'!E$2:E${0},K{1}), _xlfn.MINIFS('azure-vm-prices-base'!I$2:I${0}, 'azure-vm-prices-base'!A$2:A${0},\">=\"&E{1}*(100-{2})/100, 'azure-vm-prices-base'!B$2:B${0},\">=\"&F{1}*(100-{2})/100, 'azure-vm-prices-base'!D$2:D${0},J{1}, 'azure-vm-prices-base'!E$2:E${0},K{1})),\"\")"

formulaVM1YNamePattern      ="=IF(O{1}=\"YES\",VLOOKUP("+xls.alphabet[firstColumnCalculations + 3]+"{1} & J{1} & K{1},'azure-vm-prices-1Y'!G$2:H${0}, 2, 0),\"\")"
formulaVM1YMinPricePattern  ="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${0},  'azure-vm-prices-1Y'!A$2:A${0},\">=\"&E{1}*(100-{2})/100,     'azure-vm-prices-1Y'!B$2:B${0},\">=\"&F{1}*(100-{2})/100,   'azure-vm-prices-1Y'!D$2:D${0},J{1},   'azure-vm-prices-1Y'!E$2:E${0},K{1}),   _xlfn.MINIFS('azure-vm-prices-1Y'!I$2:I${0},   'azure-vm-prices-1Y'!A$2:A${0},\">=\"&E{1}*(100-{2})/100,   'azure-vm-prices-1Y'!B$2:B${0},\">=\"&F{1}*(100-{2})/100,   'azure-vm-prices-1Y'!D$2:D${0},J{1},   'azure-vm-prices-1Y'!E$2:E${0},K{1})),\"\")"
formulaVM3YNamePattern      ="=IF(O{1}=\"YES\",VLOOKUP("+xls.alphabet[firstColumnCalculations + 5]+"{1} & J{1} & K{1},'azure-vm-prices-3Y'!G$2:H${0}, 2, 0),\"\")"
formulaVM3YMinPricePattern  ="=IF(O{1}=\"YES\", IF(N{1}=\"YES\" ,   _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${0},  'azure-vm-prices-3Y'!A$2:A${0},\">=\"&E{1}*(100-{2})/100,     'azure-vm-prices-3Y'!B$2:B${0},\">=\"&F{1}*(100-{2})/100,   'azure-vm-prices-3Y'!D$2:D${0},J{1},   'azure-vm-prices-3Y'!E$2:E${0},K{1}),   _xlfn.MINIFS('azure-vm-prices-3Y'!I$2:I${0},   'azure-vm-prices-3Y'!A$2:A${0},\">=\"&E{1}*(100-{2})/100,   'azure-vm-prices-3Y'!B$2:B${0},\">=\"&F{1}*(100-{2})/100,   'azure-vm-prices-3Y'!D$2:D${0},J{1},   'azure-vm-prices-3Y'!E$2:E${0},K{1})),\"\")"
formulaVMYearPAYGPattern="=IF(O{0}=\"YES\",M{0}*Q{0}*12,\"\")"
formulaVMYear1YRIPattern="=IF(O{0}=\"YES\",S{0}*8760,\"\")"
formulaVMYear3YRIPattern="=IF(O{0}=\"YES\",U{0}*8760,\"\")"
formulaBestPricePattern ="=IF(O{0}=\"YES\",IF({1}=\"YES\", MIN(V{0}:X{0}), V{0}),\"\")"

#STANDARD DISK FORMULAS
formulaDiskS4Pattern ="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}<'azure-standard-disk-prices'!B2),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS6Pattern ="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B2,G{0}<'azure-standard-disk-prices'!B3),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS10Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B3,G{0}<'azure-standard-disk-prices'!B4),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS15Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B4,G{0}<'azure-standard-disk-prices'!B5),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS20Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B5,G{0}<'azure-standard-disk-prices'!B6),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS30Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B6,G{0}<'azure-standard-disk-prices'!B7),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS40Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B7,G{0}<'azure-standard-disk-prices'!B8),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskS50Pattern="=IF(AND(H{0}=\"STANDARD\",O{0}=\"YES\",G{0}>'azure-standard-disk-prices'!B8,G{0}<'azure-standard-disk-prices'!B9),1+IF(L{0}=\"YES\",1),\"\")"
#PREMIUM DISK FORMULAS
formulaDiskP4Pattern ="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}<'azure-premium-disk-prices'!B2),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP6Pattern ="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B2,G{0}<'azure-premium-disk-prices'!B3),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP10Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B3,G{0}<'azure-premium-disk-prices'!B4),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP15Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B4,G{0}<'azure-premium-disk-prices'!B5),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP20Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B5,G{0}<'azure-premium-disk-prices'!B6),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP30Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B6,G{0}<'azure-premium-disk-prices'!B7),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP40Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B7,G{0}<'azure-premium-disk-prices'!B8),1+IF(L{0}=\"YES\",1),\"\")"
formulaDiskP50Pattern="=IF(AND(H{0}=\"PREMIUM\",O{0}=\"YES\",G{0}>'azure-premium-disk-prices'!B8,G{0}<'azure-premium-disk-prices'!B9),1+IF(L{0}=\"YES\",1),\"\")"
	
#SET FORMAT FOR ALL DATA OK
customerVMDataExcelTab.conditional_format('O1:O'+str(xls.rowsForVMInput+1), {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '\"YES\"',
                                    'format':   dataOKFormat})

#CREATE COST TOTALS
customerVMDataExcelTab.merge_range('A8:B8', 'YEAR TOTALS (EUROS)', selectHeaderStyle)
customerVMDataExcelTab.write(8, 0, 'COMPUTE', selectHeaderStyle)
customerVMDataExcelTab.write(9, 0, 'STORAGE', selectHeaderStyle)
customerVMDataExcelTab.write(10,0, 'ASR',     selectHeaderStyle)
customerVMDataExcelTab.write(11, 0,'TOTAL',   selectHeaderStyle)
customerVMDataExcelTab.write_formula(8, 1, formulaTotalComputeCost, selecBody)		  
customerVMDataExcelTab.write_formula(9, 1, formulaTotalDiskCost, selecBody)
customerVMDataExcelTab.write_formula(10, 1, formulaTotalASRCost, selecBody)
customerVMDataExcelTab.write_formula(11, 1, formulaTotalCost, selectHeaderStyle)	

#CREATE DATA DISK TOTALS
customerVMDataExcelTab.merge_range('A14:C14', 'DATA DISK SUMMARY', selectHeaderStyle)
customerVMDataExcelTab.write(14, 0, 'DISK SIZE', selectHeaderStyle)
customerVMDataExcelTab.write(14, 1, 'COUNT', selectHeaderStyle)
customerVMDataExcelTab.write(14, 2, 'PRICE', selectHeaderStyle)
firstDiskTotalRow = 15
for index in range(0, len(priceReaderManagedDisk.standardDiskSizes)):
	customerVMDataExcelTab.write(firstDiskTotalRow + index, 0 , priceReaderManagedDisk.standardDiskSizes[index], selectHeaderStyle)

for index in range(0, len(priceReaderManagedDisk.premiumDiskSizes)):
	customerVMDataExcelTab.write(firstDiskTotalRow + len(priceReaderManagedDisk.standardDiskSizes) + index, 0 , priceReaderManagedDisk.premiumDiskSizes[index], selectHeaderStyle)


#CREATE OS DISK TOTALS
customerVMDataExcelTab.merge_range('A33:C33', 'DATA DISK SUMMARY', selectHeaderStyle)
customerVMDataExcelTab.write(14, 0, 'DISK SIZE', selectHeaderStyle)
customerVMDataExcelTab.write(14, 1, 'COUNT', selectHeaderStyle)
customerVMDataExcelTab.write(14, 2, 'PRICE', selectHeaderStyle)

#COLUMNS FOR REMAINING ITEMS

customerVMDataExcelTab.write(0, firstColumnCalculations + 27, 'OS DISK STANDARD' , selectHeaderStyle)
customerVMDataExcelTab.write(0, firstColumnCalculations + 28, 'OS DISK PREMIUM'  , selectHeaderStyle)

#FORMULAS AND STYLE FOR CALCULATIONS
for rowIndex in range(1,xls.rowsForVMInput):		
	formulaCheckAllInputs =formulaCheckAllInputsPattern.format( rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations - 1, formulaCheckAllInputs,   inputBodyStyle)	

	formulaVMBaseName  =formulaVMBaseNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 0, formulaVMBaseName,   selecBody)		
	
	formulaVMBaseMinPrice=formulaVMBaseMinPricePattern.format(numVmSizes, rowIndex+1, perfGainValueCell)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 1, formulaVMBaseMinPrice, selecBody)
	
	formulaVM1YName  =formulaVM1YNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 2, formulaVM1YName,   selecBody)		
	
	formulaVM1YMinPrice=formulaVM1YMinPricePattern.format(numVmSizes, rowIndex+1, perfGainValueCell)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 3, formulaVM1YMinPrice, selecBody)

	formulaVM3YName  =formulaVM3YNamePattern.format(  numVmSizes+1, rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 4, formulaVM3YName,   selecBody)		
	
	formulaVM3YMinPrice=formulaVM3YMinPricePattern.format(numVmSizes, rowIndex+1, perfGainValueCell)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 5, formulaVM3YMinPrice, selecBody)

	formulaVMYearPAYG=formulaVMYearPAYGPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 6, formulaVMYearPAYG, selecBody)

	formulaVMYear1YRI=formulaVMYear1YRIPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 7, formulaVMYear1YRI, selecBody)

	formulaVMYear3YRI=formulaVMYear3YRIPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 8, formulaVMYear3YRI, selecBody)
	
	formulaBestPrice=formulaBestPricePattern.format(rowIndex+1, resInstValueCell)
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

#COUNT OS STANDARD
	customerVMDataExcelTab.set_column(42, 43, 18) 		
	formulaDiskOSStandard=formulaDiskOSStandardPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 27, formulaDiskOSStandard, selecBody)
#COUNT OS PREMIUM
	formulaDiskOSPremium=formulaDiskOSPremiumPattern.format(rowIndex+1)
	customerVMDataExcelTab.write_formula(rowIndex, firstColumnCalculations + 28, formulaDiskOSPremium, selecBody)

	
############ OTHER TABS	
#PUT DATA FROM APIs
	#VM PAYG
azureVMDataBaseExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMDataBaseExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(2, 2, 12) 
azureVMDataBaseExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(3, 5, 9)
azureVMDataBaseExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMDataBaseExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMDataBaseExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(6, 6, 0) 
azureVMDataBaseExcelTab.set_column(7, 7, 20)
azureVMDataBaseExcelTab.write(0, 7, 'VM SIZE NAME', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(8, 8, 0)

	#VM 1Y
azureVMData1YExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMData1YExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMData1YExcelTab.set_column(2, 2, 12) 
azureVMData1YExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMData1YExcelTab.set_column(3, 5, 9)
azureVMData1YExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMData1YExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMData1YExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMData1YExcelTab.set_column(6, 6, 0) 
azureVMData1YExcelTab.set_column(7, 7, 20)
azureVMData1YExcelTab.write(0, 7, 'VM SIZE NAME', inputHeaderStyle)
azureVMData1YExcelTab.set_column(8, 8, 0) 

	#VM 3Y
azureVMData3YExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMData3YExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMData3YExcelTab.set_column(2, 2, 12) 
azureVMData3YExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMData3YExcelTab.set_column(3, 5, 9)
azureVMData3YExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMData3YExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMData3YExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMData3YExcelTab.set_column(6, 6, 0) 
azureVMData3YExcelTab.set_column(7, 7, 20)
azureVMData3YExcelTab.write(0, 7, 'VM SIZE NAME', inputHeaderStyle)
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
	azureVMDataBaseExcelTab.write_number(currentLineBase, 0, cpus,      inputBodyStyle)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 1, mem,       inputBodyStyle)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 2, priceBase, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 3, SAP, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 4, GPU, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 5, isBurstable, inputBodyStyle)
	auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLineBase+1)
	azureVMDataBaseExcelTab.write_formula(currentLineBase, 6, auxFormula, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 7, name, inputBodyStyle)
	
	noBurstablePriceBase = priceBase
	noBurstablePrice1Y = price1Y
	noBurstablePrice3Y = price3Y
	if isBurstable == "YES":
		noBurstablePriceBase += 10000
		noBurstablePrice1Y += 10000
		noBurstablePrice3Y += 10000
		
	azureVMDataBaseExcelTab.write(currentLineBase, 8, noBurstablePriceBase, inputBodyStyle)	
	
	currentLineBase += 1	

	if price1Y != 1000000:
		price1Y=float(price1Y)
		azureVMData1YExcelTab.write_number(currentLine1Y, 0, cpus,      inputBodyStyle)
		azureVMData1YExcelTab.write_number(currentLine1Y, 1, mem,       inputBodyStyle)
		azureVMData1YExcelTab.write_number(currentLine1Y, 2, price1Y, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 3, SAP, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 4, GPU, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 5, isBurstable, inputBodyStyle)
		auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLine1Y+1)
		azureVMData1YExcelTab.write_formula(currentLine1Y, 6, auxFormula, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 7, name, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 8, noBurstablePrice1Y, inputBodyStyle)	
		currentLine1Y += 1		
	
	if price3Y != 1000000:
		price3Y=float(price3Y)
		azureVMData3YExcelTab.write_number(currentLine3Y, 0, cpus,      inputBodyStyle)
		azureVMData3YExcelTab.write_number(currentLine3Y, 1, mem,       inputBodyStyle)
		azureVMData3YExcelTab.write_number(currentLine3Y, 2, price3Y, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 3, SAP, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 4, GPU, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 5, isBurstable, inputBodyStyle)
		auxFormula='=CONCATENATE(C{0},D{0},E{0})'.format(currentLine3Y+1)
		azureVMData3YExcelTab.write_formula(currentLine3Y, 6, auxFormula, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 7, name, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 8, noBurstablePrice3Y, inputBodyStyle)	
		currentLine3Y += 1		

#ASR
azureASRExcelTab.write(0, 0, 'ASR Azure to Azure', inputHeaderStyle)
azureASRExcelTab.set_column(0, 0, 25)
azureASRExcelTab.write_number(1, 0, 25,      inputBodyStyle)
for size in siteRecoveryPriceMatrix:
	priceBase = siteRecoveryPriceMatrix[size]
	azureASRExcelTab.write_number(1, 0, priceBase,      inputBodyStyle)

#PREMIUM STORAGE
azurePremiumDiskExcelTab.write(0, 0, 'Disk Size', inputHeaderStyle)
azurePremiumDiskExcelTab.set_column(0, 0, 25)
azurePremiumDiskExcelTab.set_column(1, 2, 15)
azurePremiumDiskExcelTab.write(0, 1, 'Capacity', inputHeaderStyle)
azurePremiumDiskExcelTab.write(0, 2, 'Cost', inputHeaderStyle)

currentLine = 1
for size in  sorted(premiumDiskPriceMatrix):
	priceBase = premiumDiskPriceMatrix[size]['price']
	name = premiumDiskPriceMatrix[size]['name']
	
	azurePremiumDiskExcelTab.write(currentLine, 0, name,      inputBodyStyle)
	azurePremiumDiskExcelTab.write(currentLine, 1, size,  inputBodyStyle)
	azurePremiumDiskExcelTab.write(currentLine, 2, priceBase, inputBodyStyle)
	currentLine += 1	
		
#STANDARD STORAGE 
azureStandardDiskExcelTab.write(0, 0, 'Disk Size', inputHeaderStyle)
azureStandardDiskExcelTab.set_column(0, 0, 25)
azureStandardDiskExcelTab.set_column(1, 2, 15)
azureStandardDiskExcelTab.write(0, 1, 'Capacity', inputHeaderStyle)
azureStandardDiskExcelTab.write(0, 2, 'Cost', inputHeaderStyle)

currentLine = 1
for size in  sorted(standardDiskPriceMatrix):
	priceBase = standardDiskPriceMatrix[size]['price']
	name = standardDiskPriceMatrix[size]['name']
	
	azureStandardDiskExcelTab.write(currentLine, 0, name,      inputBodyStyle)
	azureStandardDiskExcelTab.write(currentLine, 1, size,  inputBodyStyle)
	azureStandardDiskExcelTab.write(currentLine, 2, priceBase, inputBodyStyle)
	currentLine += 1	

workbook.close()    