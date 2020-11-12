#!/usr/bin/python3

import datetime
import xlsxwriter
import sys
import os

from xlsStructure import xlsStructure as xls
import priceReaderCompute
import priceReaderManagedDisk
import priceReaderSiteRecovery

workbookNamePattern = '/mnt/c/Users/segonza/Desktop/Azure-Quote-Tool-{}.xlsx'
installation_dir = '/home/sergio/azure-pricer/' 
regions=['germany-north', 'germany-west-central', 'south-africa-north', 'south-africa-west', 'switzerland-north', 'switzerland-west', 'uae-central', 'uae-north', 'asia-pacific-east', 'asia-pacific-southeast', 'australia-central', 'australia-central-2', 'australia-east','australia-southeast', 'brazil-south', 'canada-central', 'canada-east', 'central-india', 'europe-north', 'europe-west', 'france-central', 'france-south', 'germany-central', 'germany-northeast', 'japan-east', 'japan-west', 'korea-central', 'korea-south', 'south-india', 'united-kingdom-south', 'united-kingdom-west', 'us-central', 'us-east', 'us-east-2', 'usgov-arizona', 'usgov-iowa', 'usgov-texas', 'usgov-virginia', 'us-north-central', 'us-south-central', 'us-west', 'us-west-2', 'us-west-central', 'west-india', 'norway-east', 'norway-west']

regions.sort()

numVmSizesCheck   = 6900
numASRSkusCheck   = 33
numPremDisksCheck = 276
numStanDisksCheck = 276

today = datetime.date.today().strftime('%d%m%y')
workbookFile = workbookNamePattern.format(today)

if len(sys.argv) > 1:
	workbookFile=sys.argv[1]
if len(sys.argv) > 2:
	installation_dir=sys.argv[2] 
	
#KEY CELLS
perfGainValueCell=xls.getAssumptionValueCell('PERF')
modeCell=xls.getAssumptionValueCell('MODE')
#KEY COLUMN INDEX
VMNameColumn=      xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['VM NAME']['index']))
CPUColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['CPUs']['index']))
memColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['Mem(GB)']['index']))
dataDiskSizeColumn=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['DATA STORAGE']['index']))
dataDiskTypeColumn=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['DATA STORAGE TYPE']['index']))
osDiskSizeColumn=  xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['OS DISK']['index']))
regionColumn=      xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['REGION']['index']))
licenseColumn=     xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['LICENSE']['index']))
SAPColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['SAP']['index']))
GPUColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['GPU']['index']))
ASRColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ASR']['index']))
hoursMonthColumn  =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['HOURS/MONTH']['index']))
bSeriesColumn     =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['USE B SERIES']['index']))
reseInsColumn     =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['RESERVED INST']['index']))
dataOKColumn=      xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ALL DATA OK']['index']))

ssdRequiredCheckColumn=xls.getColumnLetterFromIndex(xls.ssdRequiredCheckColumn['firstColumnIndex'])
asrPriceColumn=xls.getColumnLetterFromIndex(xls.asrPriceColumn['firstColumnIndex'])

firstCalculationColumnIndex=xls.VMCalculationColumns['firstColumnIndex']

columnPAYGVMSize =     xls.getVMCalculationColumn('BEST SIZE PAYG')
column1YResInsVMSize = xls.getVMCalculationColumn('BEST SIZE 1Y')
column3YResInsVMSize = xls.getVMCalculationColumn('BEST SIZE 3Y')
columnYearPAYGVMPrice =     xls.getVMCalculationColumn('PAYG')
columnYear1YResInsVMPrice = xls.getVMCalculationColumn('1Y RI')
columnYear3YResInsVMPrice = xls.getVMCalculationColumn('3Y RI')
columnBestYearVMPrice = xls.getVMCalculationColumn('BEST PRICE')
#KEY DATA
totalNumDisks = len(priceReaderManagedDisk.standardDiskSizes) + len(priceReaderManagedDisk.premiumDiskSizes)

#1 - GET RESOURCE PRICES
computePriceMatrix = priceReaderCompute.getPriceMatrix(regions)
siteRecoveryPriceMatrix = priceReaderSiteRecovery.getPriceMatrix(regions)
premiumDiskPriceMatrix  = priceReaderManagedDisk.getPriceMatrixPremium(regions)
standardDiskPriceMatrix = priceReaderManagedDisk.getPriceMatrixStandard(regions)

numVmSizes = len(computePriceMatrix)
numSiteRecoverySKUs = len(siteRecoveryPriceMatrix)
numPremiumDiskSKUs  = len(premiumDiskPriceMatrix)
numStandardDiskSKUs = len(standardDiskPriceMatrix)  

#3 - CHECK EVERYTHING IS OK
if numVmSizes < numVmSizesCheck:
	logger("Issue with number of VMs read: " +str(numVmSizes))
	sys.exit(1)
if numSiteRecoverySKUs < numASRSkusCheck:
	logger("Issue with number of ASR SKUs read: " +str(numSiteRecoverySKUs))
	sys.exit(1)
if numPremiumDiskSKUs  < numPremDisksCheck:
	logger("Issue with number of Premium Disks read: " +str(numPremiumDiskSKUs))
	sys.exit(1)
if numStandardDiskSKUs < numStanDisksCheck:
	logger("Issue with number of Standard Disks read: " +str(numStandardDiskSKUs))
	sys.exit(1)

#2 - CREATE WORKBOOKS
workbook = xlsxwriter.Workbook(workbookFile)

#3 - ADD TABS
introTab = workbook.add_worksheet('INTRO')
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
inputHeaderStyle.set_font_color('white')
inputBodyStyle = workbook.add_format()
inputBodyStyle.set_align('center')
inputBodyStyle.set_border(1)
inputBodyStyle.set_bg_color('#e6edf6')
selectHeaderStyle = workbook.add_format()
selectHeaderStyle.set_bold()
selectHeaderStyle.set_align('center')
selectHeaderStyle.set_border(1)
selectHeaderStyle.set_bg_color('#76933c')
dataOKFormat =  workbook.add_format()
dataOKFormat.set_bold()
dataOKFormat.set_font_color('#76933c')
dataOKFormat.set_font_size(13)
dataOKFormat.set_bg_color('#b7dee8')
noVMFormat =  workbook.add_format()
noVMFormat.set_bold()
noVMFormat.set_font_size(14)
noVMFormat.set_bg_color('#ffe9d9')

selectBodyStyle = workbook.add_format()
selectBodyStyle.set_align('center')
selectBodyStyle.set_border(1)
selectBodyStyle.set_bg_color('#d8e4bc')

########################################################
#####################PUT DISCLAIMERS####################
########################################################
introTab.insert_image('A1', 'media/slide.jpg')

########################################################
#####################PUT DATA BLOCKS####################
########################################################

#1st COLUMN WIDTH
customerVMDataExcelTab.set_column(0, 0, xls.firstColumnWidth) 	

#5 - BLOCK 1 - CREATE ASSUMPTIONS
	#CALCULATE AND PUT HEADER
firstColumnLetter=xls.getColumnLetterFromIndex(xls.assumptions['firstCellColumn'])
firstRowLetter=xls.assumptions['firstCellRow'] + 1
lastColumnLetter=xls.getColumnLetterFromIndex(xls.assumptions['firstCellColumn'] + xls.assumptions['header']['width'] - 1)
headerRange='{0}{1}:{2}{1}'.format(firstColumnLetter, firstRowLetter, lastColumnLetter)
customerVMDataExcelTab.merge_range(headerRange, xls.assumptions['header']['title'], inputHeaderStyle)

	#ASSUMPTION - TOLERANCE
category='PERF'
name=xls.assumptions['rows'][category]['name']
row=xls.assumptions['firstCellRow'] + xls.assumptions['rows'][category]['order']
defaultValue=xls.assumptions['rows'][category]['default']
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'], name, inputHeaderStyle)
customerVMDataExcelTab.write_number(row, xls.assumptions['firstCellColumn'] + 1, defaultValue, inputBodyStyle)
customerVMDataExcelTab.data_validation(perfGainValueCell, {'validate': 'integer', 'criteria': 'between',
                                  'minimum': 0, 'maximum': 100, 'input_title': 'Enter an integer:',
                                  'input_message': 'between 0 and 100 on how better % Azure perf is'})

category='MODE'
name=xls.assumptions['rows'][category]['name']
row=xls.assumptions['firstCellRow'] + xls.assumptions['rows'][category]['order']
defaultValue=xls.assumptions['rows'][category]['default']
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'], name, inputHeaderStyle)
customerVMDataExcelTab.write(row, xls.assumptions['firstCellColumn'] + 1, defaultValue, inputBodyStyle)
customerVMDataExcelTab.data_validation(modeCell, {'validate': 'list','source': xls.assumptions['rows'][category]['validationList']})

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
	
	#CHECK ALL INPUT
formulaCheckAllInputsPattern="=IF(AND({1}{0}<>\"\", {2}{0}<>\"\", {3}{0}<>\"\", {4}{0}<>\"\", {5}{0}<>\"\", {6}{0}<>\"\", {7}{0}<>\"\", {8}{0}<>\"\", {9}{0}<>\"\", {10}{0}<>\"\", {11}{0}<>\"\"),\"YES\",\"NO\")"
for rowIndex in range(1,xls.rowsForVMInput):		
	formulaCheckAllInputs =formulaCheckAllInputsPattern.format( rowIndex+1, VMNameColumn, CPUColumn, memColumn, dataDiskSizeColumn, dataDiskTypeColumn, osDiskSizeColumn, SAPColumn, GPUColumn, ASRColumn, hoursMonthColumn, bSeriesColumn )
	customerVMDataExcelTab.write_formula(rowIndex, xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ALL DATA OK']['index']), formulaCheckAllInputs,   selectBodyStyle)	

	#SET FORMAT FOR ALL DATA OK
cellRange='{0}{1}:{0}{2}'.format(dataOKColumn , xls.customerInputColumns['firstCellRow'] + 1, str(xls.rowsForVMInput+1))
customerVMDataExcelTab.conditional_format(cellRange, {'type': 'cell', 'criteria': 'equal to', 'value':    '\"YES\"', 'format':   dataOKFormat})

#6 - BLOCK 3 - CREATE VM CALCULATION COLUMNS
	#PUT HEADERS
for column in xls.VMCalculationColumns['columns']:
	#GET COLUMN DATA
	columnWidth = xls.VMCalculationColumns['columns'][column]['width']
	columnName  = xls.VMCalculationColumns['columns'][column]['alias']
	columnPositon = xls.getCalculationColumnPositionInExcel(xls.VMCalculationColumns['columns'][column]['index'])
	
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnPositon, columnPositon, columnWidth)
	#SET HEADER
	customerVMDataExcelTab.write(xls.VMCalculationColumns['firstCellRow'], columnPositon, columnName, selectHeaderStyle)

#VIRTUAL MACHINE FORMULAS
formulaVMYearPAYGPattern    ="=IF({0}{1}=\"YES\", {2}{1}*{3}{1}*12*{4}{1},\"\")"
formulaVMYear1YRIPattern    ="=IF({0}{1}=\"YES\", {2}{1}*8760*{3}{1}     ,\"\")"
formulaVMYear3YRIPattern    ="=IF({0}{1}=\"YES\", {2}{1}*8760*{3}{1}     ,\"\")"

formulaBestPricePattern     ="=IF({0}{1}=\"YES\",IF({2}{1}=\"YES\", MIN({3}{1}:{4}{1}), {3}{1}),\"\")"

formulaBestPricePattern     ="=IF({0}{1}=\"YES\",  _xlfn.SWITCH({2}{1},  \"YES ALL\",  _xlfn.MINIFS({3}{1}:{4}{1}, {3}{1}:{4}{1}, \">0\"), \"YES 1Y\",  _xlfn.MINIFS({3}{1}:{5}{1}, {3}{1}:{5}{1}, \">0\"), \"NO\", {3}{1} ), \"\")"

formulaVMBaseNamePattern    ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"OK\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-base'!J$2:K${5}, 2, 0), VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"??\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-base'!J$2:K${5}, 2, 0)), \"\")"
formulaVM1YNamePattern      ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"OK\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-1Y'!J$2:K${5}  , 2, 0), VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"??\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-1Y'!J$2:K${5}, 2, 0)),   \"\")"
formulaVM3YNamePattern      ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"OK\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-3Y'!J$2:K${5}  , 2, 0), VLOOKUP({2}{1} & {6}{1} & {4}{1} & \"??\" & IF(OR(L{1}=\"PREMIUM\", ISNUMBER(SEARCH(\"P\", M{1} ) ) ), \"SDD\", \"???\" ),  'azure-vm-prices-3Y'!J$2:K${5}, 2, 0)),   \"\")"

formulaVMBaseMinPricePattern="=IF({0}{1}=\"YES\", IF( B3=\"CPU+MEM\" , _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${2}, 'azure-vm-prices-base'!A$2:A${2}, \">=\"&{3}{1}*(100-{4})/100, 'azure-vm-prices-base'!B$2:B${2}, \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-base'!G$2:G${2},{6}{1}, 'azure-vm-prices-base'!H$2:H${2}, {7}{1}, 'azure-vm-prices-base'!D$2:D${2}, IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-base'!E$2:E${2}, IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-base'!F$2:F${2}, IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-base'!I$2:I${2}, IF({11}{1}=\"YES\", \"YES\", \"*\") ), _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${2}, 'azure-vm-prices-base'!B$2:B${2}, \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-base'!G$2:G${2},{6}{1}, 'azure-vm-prices-base'!H$2:H${2}, {7}{1}, 'azure-vm-prices-base'!D$2:D${2}, IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-base'!E$2:E${2}, IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-base'!F$2:F${2}, IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-base'!I$2:I${2}, IF({11}{1}=\"YES\", \"YES\", \"*\") ) ) , \"\")"
formulaVM1YMinPricePattern=  "=IF({0}{1}=\"YES\", IF( B3=\"CPU+MEM\" , _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${2},   'azure-vm-prices-1Y'!A$2:A${2},   \">=\"&{3}{1}*(100-{4})/100, 'azure-vm-prices-1Y'!B$2:B${2},   \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-1Y'!G$2:G${2},{6}{1},   'azure-vm-prices-1Y'!H$2:H${2},   {7}{1}, 'azure-vm-prices-1Y'!D$2:D${2},   IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-1Y'!E$2:E${2},   IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-1Y'!F$2:F${2},   IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-1Y'!I$2:I${2},   IF({11}{1}=\"YES\", \"YES\", \"*\") ), _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${2}, 'azure-vm-prices-1Y'!B$2:B${2}, \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-1Y'!G$2:G${2},{6}{1}, 'azure-vm-prices-1Y'!H$2:H${2}, {7}{1}, 'azure-vm-prices-1Y'!D$2:D${2}, IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-1Y'!E$2:E${2}, IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-1Y'!F$2:F${2}, IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-1Y'!I$2:I${2}, IF({11}{1}=\"YES\", \"YES\", \"*\") ) ) , \"\")"
formulaVM3YMinPricePattern=  "=IF({0}{1}=\"YES\", IF( B3=\"CPU+MEM\" , _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${2},   'azure-vm-prices-3Y'!A$2:A${2},   \">=\"&{3}{1}*(100-{4})/100, 'azure-vm-prices-3Y'!B$2:B${2},   \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-3Y'!G$2:G${2},{6}{1},   'azure-vm-prices-3Y'!H$2:H${2},   {7}{1}, 'azure-vm-prices-3Y'!D$2:D${2},   IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-3Y'!E$2:E${2},   IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-3Y'!F$2:F${2},   IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-3Y'!I$2:I${2},   IF({11}{1}=\"YES\", \"YES\", \"*\") ), _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${2}, 'azure-vm-prices-3Y'!B$2:B${2}, \">=\"&{5}{1}*(100-{4})/100, 'azure-vm-prices-3Y'!G$2:G${2},{6}{1}, 'azure-vm-prices-3Y'!H$2:H${2}, {7}{1}, 'azure-vm-prices-3Y'!D$2:D${2}, IF({8}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-3Y'!E$2:E${2}, IF({9}{1}=\"YES\", \"YES\", \"*\"), 'azure-vm-prices-3Y'!F$2:F${2}, IF({10}{1}=\"NO\", \"NO\", \"*\"), 'azure-vm-prices-3Y'!I$2:I${2}, IF({11}{1}=\"YES\", \"YES\", \"*\") ) ) , \"\")"

#FORMULAS AND STYLE FOR CALCULATIONS
for rowIndex in range(1,xls.rowsForVMInput):
	formulaVMBaseMinPrice=formulaVMBaseMinPricePattern.format(dataOKColumn, rowIndex+1, numVmSizes, CPUColumn, perfGainValueCell, memColumn, regionColumn, licenseColumn, SAPColumn, GPUColumn, bSeriesColumn, ssdRequiredCheckColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 1, formulaVMBaseMinPrice, selectBodyStyle)
	
	formulaVM1YMinPrice=formulaVM1YMinPricePattern.format(dataOKColumn, rowIndex+1, numVmSizes, CPUColumn, perfGainValueCell, memColumn, regionColumn, licenseColumn, SAPColumn, GPUColumn, bSeriesColumn, ssdRequiredCheckColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 3, formulaVM1YMinPrice, selectBodyStyle)
	
	formulaVM3YMinPrice=formulaVM3YMinPricePattern.format(dataOKColumn, rowIndex+1, numVmSizes, CPUColumn, perfGainValueCell, memColumn, regionColumn, licenseColumn, SAPColumn, GPUColumn, bSeriesColumn, ssdRequiredCheckColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 5, formulaVM3YMinPrice, selectBodyStyle)

	formulaVMBaseName  =formulaVMBaseNamePattern.format(dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 1] , SAPColumn , GPUColumn , numVmSizes+1, regionColumn )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 0, formulaVMBaseName,   selectBodyStyle)		

	formulaVM1YName  =formulaVM1YNamePattern.format(    dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 3] , SAPColumn , GPUColumn , numVmSizes+1, regionColumn )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 2, formulaVM1YName,   selectBodyStyle)		
	
	formulaVM3YName  =formulaVM3YNamePattern.format(    dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 5] , SAPColumn , GPUColumn , numVmSizes+1, regionColumn )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 4, formulaVM3YName,   selectBodyStyle)		

	formulaVMYearPAYG=formulaVMYearPAYGPattern.format(dataOKColumn, rowIndex+1, hoursMonthColumn, xls.getVMCalculationColumn('PRICE(H) PAYG'), xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 6, formulaVMYearPAYG, selectBodyStyle)
	
	formulaVMYear1YRI=formulaVMYear1YRIPattern.format(dataOKColumn, rowIndex+1, xls.getVMCalculationColumn('PRICE(H) 1Y'), xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 7, formulaVMYear1YRI, selectBodyStyle)

	formulaVMYear3YRI=formulaVMYear3YRIPattern.format(dataOKColumn, rowIndex+1, xls.getVMCalculationColumn('PRICE(H) 3Y'), xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 8, formulaVMYear3YRI, selectBodyStyle)
	
	formulaBestPrice=formulaBestPricePattern.format(dataOKColumn, rowIndex+1, reseInsColumn, xls.getVMCalculationColumn('PAYG'), xls.getVMCalculationColumn('3Y RI'), xls.getVMCalculationColumn('1Y RI'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 9, formulaBestPrice, selectBodyStyle)	
	
#7 - BLOCK 4 - CREATE DATA DISK CALCULATION COLUMNS
	#GET  DATA
dataDiskFirstColumn = xls.managedDataDiskColumns['firstColumnIndex']
dataDiskPrefix = xls.managedDataDiskColumns['prefix']
dataDiskColumnWidth = xls.managedDataDiskColumns['width']
units = xls.managedDataDiskColumns['width']

	#STANDARD DATA DISKS
for diskIndex in range(xls.managedDataDiskColumns['firstCellRow'], len(priceReaderManagedDisk.standardDiskSizes) ):
	columnIndex = dataDiskFirstColumn + diskIndex
	diskName=priceReaderManagedDisk.standardDiskSizes[diskIndex]
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnIndex, columnIndex, dataDiskColumnWidth) 		
	#SET HEADER
	customerVMDataExcelTab.write(xls.managedDataDiskColumns['firstCellRow'], columnIndex, dataDiskPrefix+diskName, selectHeaderStyle)
	#SET FORMULAS
	diskDataIndexInDiskTab = diskIndex + 2 
	for rowIndex in range(1,xls.rowsForVMInput):
		if diskIndex == 0:
			formula ="=IF(AND({0}{1}=\"STANDARD\",{2}{1}=\"YES\",{3}{1}<='azure-standard-disk-prices'!G{4}, {3}{1}>0), ( 1+IF({5}{1}=\"YES\",1))*{6}{1},\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab, ASRColumn, xls.getCustomerDataColumn('UNITS'))
		else:
			formula ="=IF(AND({0}{1}=\"STANDARD\",{2}{1}=\"YES\",{3}{1}>'azure-standard-disk-prices'!G{4},{3}{1} <='azure-standard-disk-prices'!G{6}),( 1+IF({5}{1}=\"YES\",1))*{7}{1},\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab - 1, ASRColumn, diskDataIndexInDiskTab, xls.getCustomerDataColumn('UNITS'))
		customerVMDataExcelTab.write_formula(rowIndex, columnIndex, formula, selectBodyStyle)	
		
	#PREMIUM DATA DISKS
dataDiskFirstColumn = columnIndex + 1
for diskIndex in range(xls.managedDataDiskColumns['firstCellRow'], len(priceReaderManagedDisk.premiumDiskSizes) ):
	columnIndex = dataDiskFirstColumn + diskIndex
	diskName=priceReaderManagedDisk.premiumDiskSizes[diskIndex]
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnIndex, columnIndex, dataDiskColumnWidth) 		
	#SET HEADER
	customerVMDataExcelTab.write(xls.managedDataDiskColumns['firstCellRow'], columnIndex, dataDiskPrefix+diskName, selectHeaderStyle)
	#SET FORMULA
	diskDataIndexInDiskTab = diskIndex + 2
	for rowIndex in range(1,xls.rowsForVMInput):
		if diskIndex == 0:
			formula ="=IF(AND({0}{1}=\"PREMIUM\",{2}{1}=\"YES\",{3}{1}<'azure-premium-disk-prices'!G{4},{3}{1}>0),   ( 1+IF({5}{1}=\"YES\",1))*{6}{1},\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab, ASRColumn, xls.getCustomerDataColumn('UNITS'))
		else:
			formula ="=IF(AND({0}{1}=\"PREMIUM\",{2}{1}=\"YES\",{3}{1}>'azure-premium-disk-prices'!G{4},{3}{1}<'azure-premium-disk-prices'!G{6}),( 1+IF({5}{1}=\"YES\",1))*{7}{1},\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab - 1, ASRColumn, diskDataIndexInDiskTab, xls.getCustomerDataColumn('UNITS'))
		customerVMDataExcelTab.write_formula(rowIndex, columnIndex, formula, selectBodyStyle)	

#7 - BLOCK 5 - ASR CALCULATION COLUMNS
	#SET WIDTH
customerVMDataExcelTab.set_column(xls.ASRColumns['firstColumnIndex'], xls.ASRColumns['firstColumnIndex'], xls.ASRColumns['width'])
	#SET HEADER
customerVMDataExcelTab.write(xls.ASRColumns['firstCellRow'], xls.ASRColumns['firstColumnIndex'], xls.ASRColumns['name'] , selectHeaderStyle)
	#ROWS
formulaCheckASRPattern="=IF(AND({0}{1}=\"YES\", {2}{1}=\"YES\"),1,\"\")"	
for rowIndex in range(1,xls.rowsForVMInput):		
	formulaCheckASR=formulaCheckASRPattern.format(ASRColumn, rowIndex+1, dataOKColumn)
	customerVMDataExcelTab.write_formula(rowIndex, xls.ASRColumns['firstColumnIndex'], formulaCheckASR, selectBodyStyle)

#8 - BLOCK 6 - CREATE OS DISK CALCULATION COLUMNS
	#SET WIDTH
customerVMDataExcelTab.set_column(xls.managedS4OSDiskColumn['firstColumnIndex'],  xls.managedS4OSDiskColumn['firstColumnIndex'],  xls.managedS4OSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedP4OSDiskColumn['firstColumnIndex'],  xls.managedP4OSDiskColumn['firstColumnIndex'],  xls.managedP4OSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedS6OSDiskColumn['firstColumnIndex'],  xls.managedS6OSDiskColumn['firstColumnIndex'],  xls.managedS6OSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedP6OSDiskColumn['firstColumnIndex'],  xls.managedP6OSDiskColumn['firstColumnIndex'],  xls.managedP6OSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedS10OSDiskColumn['firstColumnIndex'], xls.managedS10OSDiskColumn['firstColumnIndex'], xls.managedS10OSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedP10OSDiskColumn['firstColumnIndex'], xls.managedP10OSDiskColumn['firstColumnIndex'], xls.managedP10OSDiskColumn['width']) 
	#SET HEADER
customerVMDataExcelTab.write(xls.managedS4OSDiskColumn['firstCellRow'],  xls.managedS4OSDiskColumn['firstColumnIndex'],  xls.managedS4OSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedP4OSDiskColumn['firstCellRow'],  xls.managedP4OSDiskColumn['firstColumnIndex'],  xls.managedP4OSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedS6OSDiskColumn['firstCellRow'],  xls.managedS6OSDiskColumn['firstColumnIndex'],  xls.managedS6OSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedP6OSDiskColumn['firstCellRow'],  xls.managedP6OSDiskColumn['firstColumnIndex'],  xls.managedP6OSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedS10OSDiskColumn['firstCellRow'], xls.managedS10OSDiskColumn['firstColumnIndex'], xls.managedS10OSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedP10OSDiskColumn['firstCellRow'], xls.managedP10OSDiskColumn['firstColumnIndex'], xls.managedP10OSDiskColumn['name'] , selectHeaderStyle)

	#ROWS
formulaDiskOSS4Pattern ="=IF(AND({0}{1}=\"S4\",  {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"
formulaDiskOSP4Pattern ="=IF(AND({0}{1}=\"P4\",  {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"
formulaDiskOSS6Pattern ="=IF(AND({0}{1}=\"S6\",  {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"
formulaDiskOSP6Pattern ="=IF(AND({0}{1}=\"P6\",  {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"
formulaDiskOSS10Pattern="=IF(AND({0}{1}=\"S10\", {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"
formulaDiskOSP10Pattern="=IF(AND({0}{1}=\"P10\", {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1)*{4}{1} ,\"\")"

for rowIndex in range(1,xls.rowsForVMInput):	
	#COUNT OS STANDARD
	formulaDiskOSStandard=formulaDiskOSS4Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedS4OSDiskColumn['firstColumnIndex'], formulaDiskOSStandard, selectBodyStyle)
	
	formulaDiskOSPremium =formulaDiskOSP4Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedP4OSDiskColumn['firstColumnIndex'] , formulaDiskOSPremium, selectBodyStyle)
	
	formulaDiskOSStandard=formulaDiskOSS6Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedS6OSDiskColumn['firstColumnIndex'], formulaDiskOSStandard, selectBodyStyle)
	
	formulaDiskOSPremium =formulaDiskOSP6Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedP6OSDiskColumn['firstColumnIndex'] , formulaDiskOSPremium, selectBodyStyle)

	formulaDiskOSStandard=formulaDiskOSS10Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedS10OSDiskColumn['firstColumnIndex'], formulaDiskOSStandard, selectBodyStyle)
	
	formulaDiskOSPremium =formulaDiskOSP10Pattern.format(osDiskSizeColumn , rowIndex+1, dataOKColumn, ASRColumn, xls.getCustomerDataColumn('UNITS'))
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedP10OSDiskColumn['firstColumnIndex'] , formulaDiskOSPremium, selectBodyStyle)

#SSD CHECK
customerVMDataExcelTab.set_column(xls.ssdRequiredCheckColumn['firstColumnIndex'], xls.ssdRequiredCheckColumn['firstColumnIndex'], xls.ssdRequiredCheckColumn['width']) 
customerVMDataExcelTab.write(xls.ssdRequiredCheckColumn['firstCellRow'],  xls.ssdRequiredCheckColumn['firstColumnIndex'],  xls.ssdRequiredCheckColumn['name'] , selectHeaderStyle)
formulaSSDCheckPattern="=IF({0}{1}=\"PREMIUM\", \"YES\", IF( OR( {2}{1}=\"P4\", {2}{1}=\"P6\", {2}{1}=\"P10\"  ), \"YES\", \"NO\"))"

for rowIndex in range(1,xls.rowsForVMInput):
	formulaSSDCheck=formulaSSDCheckPattern.format(dataDiskTypeColumn , rowIndex+1, osDiskSizeColumn)
	customerVMDataExcelTab.write(rowIndex,  xls.ssdRequiredCheckColumn['firstColumnIndex'],  formulaSSDCheck , selectBodyStyle)

#ASR COST
customerVMDataExcelTab.set_column(xls.asrPriceColumn['firstColumnIndex'], xls.asrPriceColumn['firstColumnIndex'], xls.asrPriceColumn['width']) 
customerVMDataExcelTab.write(xls.asrPriceColumn['firstCellRow'],          xls.asrPriceColumn['firstColumnIndex'], xls.asrPriceColumn['name'] , selectHeaderStyle)
formulaASRPricePattern="=IF({0}{1}=1, IFERROR(VLOOKUP({2}{1}, 'azure-asr-prices'!A$2:B${3} ,2,0)*{4}{1}, 0 ) ,\"\" )"

#DATA DISK COST
customerVMDataExcelTab.set_column(xls.diskPriceColumn['firstColumnIndex'], xls.diskPriceColumn['firstColumnIndex'], xls.diskPriceColumn['width']) 
customerVMDataExcelTab.write(xls.diskPriceColumn['firstCellRow'],          xls.diskPriceColumn['firstColumnIndex'], xls.diskPriceColumn['name'] , selectHeaderStyle)
formulaDataDiskPricePattern="=_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S4\")  * IF(AH{0}=\"\", 0,AH{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S6\")  * IF(AI{0}=\"\", 0,AI{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S10\") * IF(AJ{0}=\"\", 0,AJ{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S15\") * IF(AK{0}=\"\", 0,AK{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S20\") * IF(AL{0}=\"\", 0,AL{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S30\") * IF(AM{0}=\"\", 0,AM{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S40\") * IF(AN{0}=\"\", 0,AN{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S50\") * IF(AO{0}=\"\", 0,AO{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P4\")  * IF(AP{0}=\"\", 0,AP{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P6\")  * IF(AQ{0}=\"\", 0,AQ{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P10\") * IF(AR{0}=\"\", 0,AR{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P15\") * IF(AS{0}=\"\", 0,AS{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P20\") * IF(AT{0}=\"\", 0,AT{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P30\") * IF(AU{0}=\"\", 0,AU{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P40\") * IF(AV{0}=\"\", 0,AV{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P50\") * IF(AW{0}=\"\", 0,AW{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1},  'azure-premium-disk-prices'!A2:A{1},  \"=\" & H{0}, 'azure-premium-disk-prices'!B2:B{1},  \"=\"&L{0}) + _xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=\"&L{0})* IF(T{0}=\"NO\", 0,1)"

#OS DISK COST
customerVMDataExcelTab.set_column(xls.osDiskPriceColumn['firstColumnIndex'], xls.osDiskPriceColumn['firstColumnIndex'], xls.osDiskPriceColumn['width']) 
customerVMDataExcelTab.write(     xls.osDiskPriceColumn['firstCellRow'],     xls.osDiskPriceColumn['firstColumnIndex'], xls.osDiskPriceColumn['name'] , selectHeaderStyle)
formulaOSDiskPricePattern="=_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S4\")  * IF(AY{0}=\"\", 0,AY{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S6\")  * IF(BA{0}=\"\", 0,BA{0}) +_xlfn.MINIFS('azure-standard-disk-prices'!D2:D{2}, 'azure-standard-disk-prices'!A2:A{2}, \"=\" & J{0}, 'azure-standard-disk-prices'!B2:B{2}, \"=S10\") * IF(BC{0}=\"\", 0,BC{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P4\")  * IF(AZ{0}=\"\", 0,AZ{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P6\")  * IF(BB{0}=\"\", 0,BB{0}) +_xlfn.MINIFS('azure-premium-disk-prices'!D2:D{1}, 'azure-premium-disk-prices'!A2:A{1}, \"=\" & J{0}, 'azure-premium-disk-prices'!B2:B{1}, \"=P10\") * IF(BD{0}=\"\", 0,BD{0}) "


#DISK AND ASR
for rowIndex in range(1,xls.rowsForVMInput):
	formulaASRPrice=formulaASRPricePattern.format(xls.getColumnLetterFromIndex(xls.ASRColumns['firstColumnIndex']) , rowIndex+1, regionColumn, len(siteRecoveryPriceMatrix)+1, xls.getCustomerDataColumn('UNITS') )
	customerVMDataExcelTab.write(rowIndex,  xls.asrPriceColumn['firstColumnIndex'],  formulaASRPrice , selectBodyStyle)

	formulaDataDiskPrice=formulaDataDiskPricePattern.format( rowIndex+1, 600, 600)
	customerVMDataExcelTab.write(rowIndex,  xls.diskPriceColumn['firstColumnIndex'],  formulaDataDiskPrice , selectBodyStyle)

	formulaOSDiskPrice=formulaOSDiskPricePattern.format( rowIndex+1, 600, 600)
	customerVMDataExcelTab.write(rowIndex,  xls.osDiskPriceColumn['firstColumnIndex'],  formulaOSDiskPrice , selectBodyStyle)
	
#9 - BLOCK 7 - DATA DISK SUMMARY
	#CALCULATE AND PUT HEADER
firstColumnLetter=xls.getColumnLetterFromIndex(xls.dataDiskSummary['firstCellColumn'])
firstRowLetter=xls.dataDiskSummary['firstCellRow'] - 1
lastColumnLetter=xls.getColumnLetterFromIndex(xls.dataDiskSummary['firstCellColumn'] + xls.dataDiskSummary['header']['width'] - 1)
headerRange='{0}{1}:{2}{1}'.format(firstColumnLetter, firstRowLetter, lastColumnLetter)
customerVMDataExcelTab.merge_range(headerRange, xls.dataDiskSummary['header']['title'], selectHeaderStyle)
	#PUT COLUMN HEADERS
for column in xls.dataDiskSummary['columns']:
	customerVMDataExcelTab.write(xls.dataDiskSummary['firstCellRow'] - 1, xls.dataDiskSummary['columns'][column]['order'] - 1, xls.dataDiskSummary['columns'][column]['name'], selectHeaderStyle)
	#STANDARD DATA DISKS 
for index in range(0, len(priceReaderManagedDisk.standardDiskSizes)):
	customerVMDataExcelTab.write(xls.dataDiskSummary['firstCellRow'] + index, 0 , priceReaderManagedDisk.standardDiskSizes[index], selectHeaderStyle)
	#PREMIUM DATA DISKS 
for index in range(0, len(priceReaderManagedDisk.premiumDiskSizes)):
	customerVMDataExcelTab.write(xls.dataDiskSummary['firstCellRow'] + len(priceReaderManagedDisk.standardDiskSizes) + index, 0 , priceReaderManagedDisk.premiumDiskSizes[index], selectHeaderStyle)
	#STANDARD DATA DISKS SUMMARY CALCULATIONS
for columnIndex in range(xls.managedDataDiskColumns['firstColumnIndex'], xls.managedDataDiskColumns['firstColumnIndex'] + len(priceReaderManagedDisk.standardDiskSizes)):
	formulaCountDisk="=SUM({0}{1}:{0}{2})".format( xls.alphabet[columnIndex], xls.managedDataDiskColumns['firstCellRow'] + 1, xls.rowsForVMInput + 1)
	currentCountRow=columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] + xls.dataDiskSummary['firstCellRow'] + 1
	currentDiskPriceRow= columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] + 2
	
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 1, formulaCountDisk, selectBodyStyle)
	#PREMIUM DATA DISKS SUMMARY CALCULATIONS
for columnIndex in range(xls.managedDataDiskColumns['firstColumnIndex']  + len(priceReaderManagedDisk.standardDiskSizes), xls.managedDataDiskColumns['firstColumnIndex'] + len(priceReaderManagedDisk.standardDiskSizes) + len(priceReaderManagedDisk.premiumDiskSizes)):
	formulaCountDisk="=SUM({0}{1}:{0}{2})".format( xls.alphabet[columnIndex], xls.managedDataDiskColumns['firstCellRow'] + 1, xls.rowsForVMInput + 1)
	currentCountRow=columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] + xls.dataDiskSummary['firstCellRow'] + 1
	currentDiskPriceRow= columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] - len(priceReaderManagedDisk.standardDiskSizes) + 2
	
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 1, formulaCountDisk, selectBodyStyle)

#10 - BLOCK 8 - OS DISK SUMMARY
	#CALCULATE AND PUT HEADER
firstColumnLetter=xls.getColumnLetterFromIndex(xls.OSDiskSummary['firstCellColumn'])
firstRowLetter=xls.OSDiskSummary['firstCellRow'] - 1
lastColumnLetter=xls.getColumnLetterFromIndex(xls.OSDiskSummary['firstCellColumn'] + xls.OSDiskSummary['header']['width'] - 1)
headerRange='{0}{1}:{2}{1}'.format(firstColumnLetter, firstRowLetter, lastColumnLetter)
customerVMDataExcelTab.merge_range(headerRange, xls.OSDiskSummary['header']['title'], selectHeaderStyle)
	#PUT COLUMN HEADERS
for column in xls.OSDiskSummary['columns']:
	customerVMDataExcelTab.write(xls.OSDiskSummary['firstCellRow'] - 1, xls.OSDiskSummary['columns'][column]['order'] - 1, xls.OSDiskSummary['columns'][column]['name'], selectHeaderStyle)
	#PUT ROWS
S4OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedS4OSDiskColumn['firstColumnIndex']], xls.managedS4OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)
P4OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedP4OSDiskColumn['firstColumnIndex']], xls.managedP4OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)

S6OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedS6OSDiskColumn['firstColumnIndex']], xls.managedS6OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)
P6OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedP6OSDiskColumn['firstColumnIndex']], xls.managedP6OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)

S10OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedS10OSDiskColumn['firstColumnIndex']], xls.managedS10OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)
P10OSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedP10OSDiskColumn['firstColumnIndex']], xls.managedP10OSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)

for row in xls.OSDiskSummary['rows']:
	currentRowIndex = xls.OSDiskSummary['firstCellRow'] - 1 + xls.OSDiskSummary['rows'][row]['order']
	if row == 'S4':
		countFormula=S4OSDiskCountFormula
	elif row == 'P4':
		countFormula=P4OSDiskCountFormula	
	elif row == 'S6':
		countFormula=S6OSDiskCountFormula	
	elif row == 'P6':
		countFormula=P6OSDiskCountFormula	
	elif row == 'S10':
		countFormula=S10OSDiskCountFormula	
	elif row == 'P10':
		countFormula=P10OSDiskCountFormula	
		
	customerVMDataExcelTab.write(currentRowIndex, xls.OSDiskSummary['firstCellColumn'], xls.OSDiskSummary['rows'][row]['name'], selectHeaderStyle)
	customerVMDataExcelTab.write_formula(currentRowIndex, xls.OSDiskSummary['firstCellColumn'] + 1 , countFormula, selectBodyStyle)
	
#11 - BLOCK 9 - COST SUMMARY
formulaTotalComputeCost="=SUM({0}1:{0}{1})".format(columnBestYearVMPrice, xls.rowsForVMInput + 1)
formulaTotalDiskCost="=12*( SUM(BG2:BG{0}) + SUM(BH2:BH{0}) )".format(xls.rowsForVMInput+1)
formulaTotalASRCost ="=12*SUM({0}{1}:{0}{2})".format(xls.getColumnLetterFromIndex(xls.asrPriceColumn['firstColumnIndex']), xls.ASRColumns['firstCellRow'] + 1, xls.rowsForVMInput + 1)
formulaTotalCost="=SUM(B{0}:B{1})".format(xls.costSummary['firstCellRow'] + 2, xls.costSummary['firstCellRow'] + 2 + len(xls.costSummary['rows']) - 2)
	#CALCULATE AND PUT HEADER
firstColumnLetter=xls.getColumnLetterFromIndex(xls.costSummary['firstCellColumn'])
firstRowLetter=xls.costSummary['firstCellRow'] + 1
lastColumnLetter=xls.getColumnLetterFromIndex(xls.costSummary['firstCellColumn'] + xls.costSummary['header']['width'] - 1)
headerRange='{0}{1}:{2}{1}'.format(firstColumnLetter, firstRowLetter, lastColumnLetter)
customerVMDataExcelTab.merge_range(headerRange, xls.costSummary['header']['title'], selectHeaderStyle)
	#CREATE COST TOTALS
customerVMDataExcelTab.write(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['COMPUTE']['order'], xls.costSummary['firstCellColumn'], xls.costSummary['rows']['COMPUTE']['name'], selectHeaderStyle)
customerVMDataExcelTab.write(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['STORAGE']['order'], xls.costSummary['firstCellColumn'], xls.costSummary['rows']['STORAGE']['name'], selectHeaderStyle)
customerVMDataExcelTab.write(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['ASR']['order']    , xls.costSummary['firstCellColumn'], xls.costSummary['rows']['ASR']['name']    , selectHeaderStyle)
customerVMDataExcelTab.write(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['TOTAL']['order']  , xls.costSummary['firstCellColumn'], xls.costSummary['rows']['TOTAL']['name'],   selectHeaderStyle)
customerVMDataExcelTab.write_formula(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['COMPUTE']['order'], xls.costSummary['firstCellColumn'] + 1, formulaTotalComputeCost, selectBodyStyle)		  
customerVMDataExcelTab.write_formula(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['STORAGE']['order'], xls.costSummary['firstCellColumn'] + 1, formulaTotalDiskCost, selectBodyStyle)
customerVMDataExcelTab.write_formula(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['ASR']['order']    , xls.costSummary['firstCellColumn'] + 1, formulaTotalASRCost, selectBodyStyle)
customerVMDataExcelTab.write_formula(xls.costSummary['firstCellRow'] + xls.costSummary['rows']['TOTAL']['order']  , xls.costSummary['firstCellColumn'] + 1, formulaTotalCost, selectHeaderStyle)

#11 - BLOCK 10 - CREATE VM CALCULATION COLUMNS
	#PUT HEADERS
for column in xls.computeSummaryColumns['columns']:
	#GET COLUMN DATA
	columnWidth = xls.computeSummaryColumns['columns'][column]['width']
	columnName  = xls.computeSummaryColumns['columns'][column]['alias']
	columnPositon = xls.computeSummaryColumns['firstColumnIndex'] + xls.computeSummaryColumns['columns'][column]['index']
	
	#SET WIDTH
	customerVMDataExcelTab.set_column(columnPositon, columnPositon, columnWidth)
	#SET HEADER
	customerVMDataExcelTab.write(xls.computeSummaryColumns['firstCellRow'], columnPositon, columnName, selectHeaderStyle)

	formulaCheapestSizePattern ="=IF({0}{1}={2}{1}, {3}{1}, IF({0}{1}={4}{1},{5}{1},{6}{1}))"	
	formulaCheapestPricePattern="=IF({0}{1}=0, \"NO VM FOUND\", {0}{1})"
	formulaCheapestModelPattern="= IF({2}{1}=\"\" ,\"\",IF({0}{1}={2}{1}, \"PAYG\", IF({0}{1}={3}{1},\"1Y RI\",\"3Y RI\")))"	 
	
	#FORMULAS AND STYLE FOR CALCULATIONS
	for rowIndex in range(1,xls.rowsForVMInput):
		if   column == 'CHEAPEST SIZE':
			formula=formulaCheapestSizePattern.format(columnBestYearVMPrice, rowIndex+1, columnYearPAYGVMPrice, columnPAYGVMSize, columnYear1YResInsVMPrice, column1YResInsVMSize, column3YResInsVMSize)
		elif column == 'CHEAPEST PRICE':
			formula=formulaCheapestPricePattern.format(columnBestYearVMPrice, rowIndex+1)
		elif column == 'CHEAPEST MODEL':
			formula=formulaCheapestModelPattern.format(columnBestYearVMPrice, rowIndex+1, columnYearPAYGVMPrice, columnYear1YResInsVMPrice)		
 
		customerVMDataExcelTab.write_formula(rowIndex, columnPositon, formula, selectBodyStyle)

#IF NO VM COULD BE FOUND PUT A WARNING
	#SET FORMAT FOR ALL DATA OK
cellRange='{0}{1}:{0}{2}'.format("U" , xls.customerInputColumns['firstCellRow'] + 1, str(xls.rowsForVMInput+1))
customerVMDataExcelTab.conditional_format(cellRange, {'type': 'cell', 'criteria': 'equal to', 'value':    '\"NO VM FOUND\"', 'format':   noVMFormat})

		
######### PUT A LIST OF REGIONS ############################
for index in range(0, len(regions)):
	customerVMDataExcelTab.write(index, xls.regionListColumn, regions[index])

####################################################################
############################ OTHER TABS	############################
####################################################################
#PUT DATA FROM APIs
	#VM PAYG
azureVMDataBaseExcelTab.set_column(0, 0, 5) 
azureVMDataBaseExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(1, 1, 9) 
azureVMDataBaseExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(2, 2, 12) 
azureVMDataBaseExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(3, 3, 4)
azureVMDataBaseExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(4, 4, 5)
azureVMDataBaseExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(5, 5, 9)
azureVMDataBaseExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(6, 6, 21)
azureVMDataBaseExcelTab.write(0, 6, 'Region', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(7, 7, 9)
azureVMDataBaseExcelTab.write(0, 7, 'License', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(8, 8, 4)
azureVMDataBaseExcelTab.write(0, 8, 'SSD', inputHeaderStyle)
azureVMDataBaseExcelTab.set_column(9, 9, 0)
azureVMDataBaseExcelTab.set_column(10, 10, 17)
azureVMDataBaseExcelTab.write(0, 10, 'VM SIZE NAME', inputHeaderStyle)

	#VM 1Y
azureVMData1YExcelTab.set_column(0, 0, 5) 
azureVMData1YExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMData1YExcelTab.set_column(1, 1, 9) 
azureVMData1YExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMData1YExcelTab.set_column(2, 2, 12) 
azureVMData1YExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMData1YExcelTab.set_column(3, 3, 4)
azureVMData1YExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMData1YExcelTab.set_column(4, 4, 5)
azureVMData1YExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMData1YExcelTab.set_column(5, 5, 9)
azureVMData1YExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMData1YExcelTab.set_column(6, 6, 21)
azureVMData1YExcelTab.write(0, 6, 'Region', inputHeaderStyle)
azureVMData1YExcelTab.set_column(7, 7, 9)
azureVMData1YExcelTab.write(0, 7, 'License', inputHeaderStyle)
azureVMData1YExcelTab.set_column(8, 8, 4)
azureVMData1YExcelTab.write(0, 8, 'SSD', inputHeaderStyle)
azureVMData1YExcelTab.set_column(9, 9, 0)
azureVMData1YExcelTab.set_column(10, 10, 17)
azureVMData1YExcelTab.write(0, 10, 'VM SIZE NAME', inputHeaderStyle)

	#VM 3Y
azureVMData3YExcelTab.set_column(0, 0, 5) 
azureVMData3YExcelTab.write(0, 0, 'CPUs', inputHeaderStyle)
azureVMData3YExcelTab.set_column(1, 1, 9) 
azureVMData3YExcelTab.write(0, 1, 'Mem(GB)', inputHeaderStyle)
azureVMData3YExcelTab.set_column(2, 2, 12) 
azureVMData3YExcelTab.write(0, 2, 'Price/Hour', inputHeaderStyle)
azureVMData3YExcelTab.set_column(3, 3, 4)
azureVMData3YExcelTab.write(0, 3, 'SAP', inputHeaderStyle)
azureVMData3YExcelTab.set_column(4, 4, 5)
azureVMData3YExcelTab.write(0, 4, 'GPU', inputHeaderStyle)
azureVMData3YExcelTab.set_column(5, 5, 9)
azureVMData3YExcelTab.write(0, 5, 'Burstable', inputHeaderStyle)
azureVMData3YExcelTab.set_column(6, 6, 21)
azureVMData3YExcelTab.write(0, 6, 'Region', inputHeaderStyle)
azureVMData3YExcelTab.set_column(7, 7, 9)
azureVMData3YExcelTab.write(0, 7, 'License', inputHeaderStyle)
azureVMData3YExcelTab.set_column(8, 8, 4)
azureVMData3YExcelTab.write(0, 8, 'SSD', inputHeaderStyle)
azureVMData3YExcelTab.set_column(9, 9, 0)
azureVMData3YExcelTab.set_column(10, 10, 17)
azureVMData3YExcelTab.write(0, 10, 'VM SIZE NAME', inputHeaderStyle)

currentLineBase = 1
currentLine1Y = 1
currentLine3Y = 1

	#DUMP API DATA
for size in sorted(computePriceMatrix, reverse=True):

	cpus = computePriceMatrix[size]['cpu']
	mem  = computePriceMatrix[size]['ram']
	priceBase = computePriceMatrix[size]['payg']
	price1Y=computePriceMatrix[size]['1y']
	price3Y=computePriceMatrix[size]['3y']	
	name=size.split("-")[0]+"-"+size.split("-")[1]
	SAP=computePriceMatrix[size]['sap']
	GPU=computePriceMatrix[size]['gpu']
	isBurstable=computePriceMatrix[size]['burstable']
	region=computePriceMatrix[size]['region']
	license=computePriceMatrix[size]['os']	
	ssd=computePriceMatrix[size]['ssd']	
	
	#WRITE TO THE BASIC
	priceBase=float(priceBase)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 0, cpus,      inputBodyStyle)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 1, mem,       inputBodyStyle)
	azureVMDataBaseExcelTab.write_number(currentLineBase, 2, priceBase, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 3, SAP, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 4, GPU, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 5, isBurstable, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 6, region, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 7, license, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 8, ssd, inputBodyStyle)	
	auxFormula='=CONCATENATE(C{0},G{0},E{0},  IF(D{0}=\"YES\",\"OK\", \"KO\" ),  IF(I{0}=\"YES\",\"SDD\", \"HDD\" )  )'.format(currentLineBase+1)
	azureVMDataBaseExcelTab.write_formula(currentLineBase, 9, auxFormula, inputBodyStyle)
	azureVMDataBaseExcelTab.write(currentLineBase, 10, name, inputBodyStyle)
	
	currentLineBase += 1	

	if price1Y != 1000000:
		price1Y=float(price1Y)
		azureVMData1YExcelTab.write_number(currentLine1Y, 0, cpus,      inputBodyStyle)
		azureVMData1YExcelTab.write_number(currentLine1Y, 1, mem,       inputBodyStyle)
		azureVMData1YExcelTab.write_number(currentLine1Y, 2, price1Y, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 3, SAP, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 4, GPU, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 5, isBurstable, inputBodyStyle)	
		azureVMData1YExcelTab.write(currentLine1Y, 6, region, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 7, license, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 8, ssd, inputBodyStyle)	
		auxFormula='=CONCATENATE(C{0},G{0},E{0},  IF(D{0}=\"YES\",\"OK\", \"KO\" ),  IF(I{0}=\"YES\",\"SDD\", \"HDD\" )  )'.format(currentLine1Y+1)
		azureVMData1YExcelTab.write_formula(currentLine1Y, 9, auxFormula, inputBodyStyle)
		azureVMData1YExcelTab.write(currentLine1Y, 10, name, inputBodyStyle)	
		currentLine1Y += 1		
	
	if price3Y != 1000000:
		price3Y=float(price3Y)
		azureVMData3YExcelTab.write_number(currentLine3Y, 0, cpus,      inputBodyStyle)
		azureVMData3YExcelTab.write_number(currentLine3Y, 1, mem,       inputBodyStyle)
		azureVMData3YExcelTab.write_number(currentLine3Y, 2, price3Y, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 3, SAP, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 4, GPU, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 5, isBurstable, inputBodyStyle)	
		azureVMData3YExcelTab.write(currentLine3Y, 6, region, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 7, license, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 8, ssd, inputBodyStyle)	
		auxFormula='=CONCATENATE(C{0},G{0},E{0},  IF(D{0}=\"YES\",\"OK\", \"KO\" ),  IF(I{0}=\"YES\",\"SDD\", \"HDD\" )  )'.format(currentLine3Y+1)
		azureVMData3YExcelTab.write_formula(currentLine3Y, 9, auxFormula, inputBodyStyle)
		azureVMData3YExcelTab.write(currentLine3Y, 10, name, inputBodyStyle)	
		currentLine3Y += 1		

#ASR
azureASRExcelTab.write(0, 0, 'Region', inputHeaderStyle)
azureASRExcelTab.write(0, 1, 'ASR Azure to Azure', inputHeaderStyle)
azureASRExcelTab.set_column(0, 0, 20)
azureASRExcelTab.set_column(0, 1, 25)

regionIndex=1
for region in siteRecoveryPriceMatrix:
	priceBase = siteRecoveryPriceMatrix[region]
	azureASRExcelTab.write(regionIndex, 0, region,    inputBodyStyle)
	azureASRExcelTab.write_number(regionIndex, 1, priceBase, inputBodyStyle)
	regionIndex += 1
#PREMIUM STORAGE
azurePremiumDiskExcelTab.write(0, 0, 'Region', inputHeaderStyle)
azurePremiumDiskExcelTab.write(0, 1, 'Disk Size', inputHeaderStyle)
azurePremiumDiskExcelTab.set_column(0, 0, 25)
azurePremiumDiskExcelTab.set_column(1, 2, 15)
azurePremiumDiskExcelTab.write(0, 2, 'Capacity', inputHeaderStyle)
azurePremiumDiskExcelTab.write(0, 3, 'Cost', inputHeaderStyle)
azurePremiumDiskExcelTab.write(0, 5, 'DISK NAME', inputHeaderStyle)
azurePremiumDiskExcelTab.write(0, 6, 'DISK CAPACITY', inputHeaderStyle)

currentLine = 1
for size in  sorted(premiumDiskPriceMatrix):
	region = premiumDiskPriceMatrix[size]['region']
	priceBase = premiumDiskPriceMatrix[size]['price']
	name = premiumDiskPriceMatrix[size]['name']
	capacity = premiumDiskPriceMatrix[size]['size']
	
	azurePremiumDiskExcelTab.write(currentLine, 0, region,    inputBodyStyle)	
	azurePremiumDiskExcelTab.write(currentLine, 1, name,      inputBodyStyle)
	azurePremiumDiskExcelTab.write(currentLine, 2, capacity,  inputBodyStyle)
	azurePremiumDiskExcelTab.write(currentLine, 3, priceBase, inputBodyStyle)
	currentLine += 1	

currentLine=1	
for size in xls.premiumDisks:
	azurePremiumDiskExcelTab.write(currentLine, 5, size['diskName'], inputBodyStyle)
	azurePremiumDiskExcelTab.write(currentLine, 6, size['diskSize'], inputBodyStyle)	
	currentLine += 1
	
#STANDARD STORAGE 
azureStandardDiskExcelTab.write(0, 0, 'Region', inputHeaderStyle)
azureStandardDiskExcelTab.write(0, 1, 'Disk Size', inputHeaderStyle)
azureStandardDiskExcelTab.set_column(0, 0, 25)
azureStandardDiskExcelTab.set_column(1, 2, 15)
azureStandardDiskExcelTab.write(0, 2, 'Capacity', inputHeaderStyle)
azureStandardDiskExcelTab.write(0, 3, 'Cost', inputHeaderStyle)
azureStandardDiskExcelTab.write(0, 5, 'DISK NAME', inputHeaderStyle)
azureStandardDiskExcelTab.write(0, 6, 'DISK CAPACITY', inputHeaderStyle)


currentLine = 1
for size in  sorted(standardDiskPriceMatrix):
	region = standardDiskPriceMatrix[size]['region']
	priceBase = standardDiskPriceMatrix[size]['price']
	name = standardDiskPriceMatrix[size]['name']
	capacity = standardDiskPriceMatrix[size]['size']
	
	azureStandardDiskExcelTab.write(currentLine, 0, region,    inputBodyStyle)	
	azureStandardDiskExcelTab.write(currentLine, 1, name,      inputBodyStyle)
	azureStandardDiskExcelTab.write(currentLine, 2, capacity,  inputBodyStyle)
	azureStandardDiskExcelTab.write(currentLine, 3, priceBase, inputBodyStyle)
	currentLine += 1	

currentLine=1	
for size in xls.standardDisks:
	azureStandardDiskExcelTab.write(currentLine, 5, size['diskName'], inputBodyStyle)
	azureStandardDiskExcelTab.write(currentLine, 6, size['diskSize'], inputBodyStyle)	
	currentLine += 1
	
workbook.close()    
