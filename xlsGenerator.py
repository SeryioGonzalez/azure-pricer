#!/usr/bin/python3

import datetime
import xlsxwriter
import sys

from xlsStructure import xlsStructure as xls
import priceReaderCompute
import priceReaderManagedDisk
import priceReaderSiteRecovery

region = 'europe-west'

workbookNamePattern = '/mnt/c/Users/segonza/Desktop/Azure-Quote-Tool-{}.xlsx'
today = datetime.date.today().strftime('%d%m%y')
workbookFile = workbookNamePattern.format(today)

if len(sys.argv) > 1:
	workbookFile=sys.argv[1]

#KEY CELLS
perfGainValueCell=xls.getAssumptionValueCell('PERF')
#KEY COLUMN INDEX
VMNameColumn=      xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['VM NAME']['index']))
CPUColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['CPUs']['index']))
memColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['Mem(GB)']['index']))
dataDiskSizeColumn=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['DATA STORAGE']['index']))
dataDiskTypeColumn=xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['DATA STORAGE TYPE']['index']))
osDiskTypeColumn=  xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['OS STORAGE TYPE']['index']))
SAPColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['SAP']['index']))
GPUColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['GPU']['index']))
ASRColumn=         xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ASR']['index']))
hoursMonthColumn  =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['HOURS/MONTH']['index']))
bSeriesColumn     =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['USE B SERIES']['index']))
reseInsColumn     =xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['RESERVED INST']['index']))
dataOKColumn=      xls.getColumnLetterFromIndex(xls.getCustomerDataColumnPositionInExcel(xls.customerInputColumns['columns']['ALL DATA OK']['index']))

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
computePriceMatrix = priceReaderCompute.getPriceMatrix(region)
siteRecoveryPriceMatrix = priceReaderSiteRecovery.getPriceMatrix(region)
premiumDiskPriceMatrix  = priceReaderManagedDisk.getPriceMatrixPremium(region)
standardDiskPriceMatrix = priceReaderManagedDisk.getPriceMatrixStandard(region)
numVmSizes = len(computePriceMatrix)

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
	formulaCheckAllInputs =formulaCheckAllInputsPattern.format( rowIndex+1, VMNameColumn, CPUColumn, memColumn, dataDiskSizeColumn, dataDiskTypeColumn, osDiskTypeColumn, SAPColumn, GPUColumn, ASRColumn, hoursMonthColumn, bSeriesColumn )
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
formulaVMYearPAYGPattern    ="=IF({0}{1}=\"YES\",{2}{1}*{3}{1}*12,\"\")"
formulaVMYear1YRIPattern    ="=IF({0}{1}=\"YES\",{2}{1}*8760,\"\")"
formulaVMYear3YRIPattern    ="=IF({0}{1}=\"YES\",{2}{1}*8760,\"\")"
formulaBestPricePattern     ="=IF({0}{1}=\"YES\",IF({2}{1}=\"YES\", MIN({3}{1}:{4}{1}), {3}{1}),\"\")"

formulaVMBaseNamePattern    ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {4}{1} & {3}{1},'azure-vm-prices-base'!G$2:H${5}, 2, 0), VLOOKUP({2}{1} & {4}{1} & \"*\",'azure-vm-prices-base'!G$2:H${5}, 2, 0)), \"\")"
formulaVM1YNamePattern      ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {4}{1} & {3}{1},'azure-vm-prices-1Y'!G$2:H${5}  , 2, 0), VLOOKUP({2}{1} & {4}{1} & \"*\",'azure-vm-prices-1Y'!G$2:H${5}, 2, 0)),   \"\")"
formulaVM3YNamePattern      ="=IF({0}{1}=\"YES\", IF({3}{1}=\"YES\", VLOOKUP({2}{1} & {4}{1} & {3}{1},'azure-vm-prices-3Y'!G$2:H${5}  , 2, 0), VLOOKUP({2}{1} & {4}{1} & \"*\",'azure-vm-prices-3Y'!G$2:H${5}, 2, 0)),   \"\")"

formulaVMBaseMinPricePattern="=IF({0}{1}=\"YES\", IF({2}{1}=\"NO\" , IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-base'!I$2:I${3}, 'azure-vm-prices-base'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100, 'azure-vm-prices-base'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100, 'azure-vm-prices-base'!D$2:D${3},{7}{1}, 'azure-vm-prices-base'!E$2:E${3},{8}{1}), _xlfn.MINIFS('azure-vm-prices-base'!I$2:I${3}, 'azure-vm-prices-base'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100, 'azure-vm-prices-base'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100, 'azure-vm-prices-base'!E$2:E${3},{8}{1})), IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${3}, 'azure-vm-prices-base'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100, 'azure-vm-prices-base'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100, 'azure-vm-prices-base'!D$2:D${3},{7}{1}, 'azure-vm-prices-base'!E$2:E${3},{8}{1}), _xlfn.MINIFS('azure-vm-prices-base'!C$2:C${3}, 'azure-vm-prices-base'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100, 'azure-vm-prices-base'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100, 'azure-vm-prices-base'!E$2:E${3},{8}{1}))), \"\")"
formulaVM1YMinPricePattern  ="=IF({0}{1}=\"YES\", IF({2}{1}=\"NO\" , IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-1Y'!I$2:I${3},   'azure-vm-prices-1Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!D$2:D${3},{7}{1},   'azure-vm-prices-1Y'!E$2:E${3},{8}{1}),   _xlfn.MINIFS('azure-vm-prices-1Y'!I$2:I${3},   'azure-vm-prices-1Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!E$2:E${3},{8}{1})),   IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${3},   'azure-vm-prices-1Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!D$2:D${3},{7}{1},   'azure-vm-prices-1Y'!E$2:E${3},{8}{1}),   _xlfn.MINIFS('azure-vm-prices-1Y'!C$2:C${3},   'azure-vm-prices-1Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-1Y'!E$2:E${3},{8}{1}))),   \"\")"
formulaVM3YMinPricePattern  ="=IF({0}{1}=\"YES\", IF({2}{1}=\"NO\" , IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-3Y'!I$2:I${3},   'azure-vm-prices-3Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!D$2:D${3},{7}{1},   'azure-vm-prices-3Y'!E$2:E${3},{8}{1}),   _xlfn.MINIFS('azure-vm-prices-3Y'!I$2:I${3},   'azure-vm-prices-3Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!E$2:E${3},{8}{1})),   IF({7}{1}=\"YES\", _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${3},   'azure-vm-prices-3Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!D$2:D${3},{7}{1},   'azure-vm-prices-3Y'!E$2:E${3},{8}{1}),   _xlfn.MINIFS('azure-vm-prices-3Y'!C$2:C${3},   'azure-vm-prices-3Y'!A$2:A${3},\">=\"&{4}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!B$2:B${3},\">=\"&{6}{1}*(100-{5})/100,   'azure-vm-prices-3Y'!E$2:E${3},{8}{1}))),   \"\")"

#FORMULAS AND STYLE FOR CALCULATIONS
for rowIndex in range(1,xls.rowsForVMInput):
	formulaVMBaseMinPrice=formulaVMBaseMinPricePattern.format(dataOKColumn, rowIndex+1, bSeriesColumn, numVmSizes, CPUColumn, perfGainValueCell, memColumn, SAPColumn, GPUColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 1, formulaVMBaseMinPrice, selectBodyStyle)
	
	formulaVM1YMinPrice=formulaVM1YMinPricePattern.format(dataOKColumn, rowIndex+1, bSeriesColumn, numVmSizes, CPUColumn, perfGainValueCell, memColumn, SAPColumn, GPUColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 3, formulaVM1YMinPrice, selectBodyStyle)
	
	formulaVM3YMinPrice=formulaVM3YMinPricePattern.format(dataOKColumn, rowIndex+1, bSeriesColumn, numVmSizes, CPUColumn, perfGainValueCell, memColumn, SAPColumn, GPUColumn)
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 5, formulaVM3YMinPrice, selectBodyStyle)

	formulaVMBaseName  =formulaVMBaseNamePattern.format(dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 1] , SAPColumn , GPUColumn , numVmSizes+1 )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 0, formulaVMBaseName,   selectBodyStyle)		

	formulaVM1YName  =formulaVM1YNamePattern.format(    dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 3] , SAPColumn , GPUColumn , numVmSizes+1 )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 2, formulaVM1YName,   selectBodyStyle)		
	
	formulaVM3YName  =formulaVM3YNamePattern.format(    dataOKColumn, rowIndex+1, xls.alphabet[firstCalculationColumnIndex + 5] , SAPColumn , GPUColumn , numVmSizes+1 )
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 4, formulaVM3YName,   selectBodyStyle)		

	formulaVMYearPAYG=formulaVMYearPAYGPattern.format(dataOKColumn, rowIndex+1, hoursMonthColumn, xls.getVMCalculationColumn('PRICE(H) PAYG'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 6, formulaVMYearPAYG, selectBodyStyle)
	
	formulaVMYear1YRI=formulaVMYear1YRIPattern.format(dataOKColumn, rowIndex+1, xls.getVMCalculationColumn('PRICE(H) 1Y'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 7, formulaVMYear1YRI, selectBodyStyle)

	formulaVMYear3YRI=formulaVMYear3YRIPattern.format(dataOKColumn, rowIndex+1, xls.getVMCalculationColumn('PRICE(H) 3Y'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 8, formulaVMYear3YRI, selectBodyStyle)
	
	formulaBestPrice=formulaBestPricePattern.format(dataOKColumn, rowIndex+1, reseInsColumn, xls.getVMCalculationColumn('PAYG'), xls.getVMCalculationColumn('3Y RI'))
	customerVMDataExcelTab.write_formula(rowIndex, firstCalculationColumnIndex + 9, formulaBestPrice, selectBodyStyle)	
	
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
	#SET FORMULA
	diskDataIndexInDiskTab = diskIndex + 2
	for rowIndex in range(1,xls.rowsForVMInput):
		if diskIndex == 0:
			formula ="=IF(AND({0}{1}=\"STANDARD\",{2}{1}=\"YES\",{3}{1}<'azure-standard-disk-prices'!B{4}, {3}{1}>0),1+IF({5}{1}=\"YES\",1),\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab, ASRColumn)
		else:
			formula ="=IF(AND({0}{1}=\"STANDARD\",{2}{1}=\"YES\",{3}{1}>'azure-standard-disk-prices'!B{4},{3}{1}<'azure-standard-disk-prices'!B{6}),1+IF({5}{1}=\"YES\",1),\"\")".format(dataDiskTypeColumn, rowIndex + 1, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab - 1, ASRColumn, diskDataIndexInDiskTab)
		customerVMDataExcelTab.write_formula(rowIndex, columnIndex, formula, selectBodyStyle)	
		
	#PREMIUM DISKS
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
			formula ="=IF(AND({0}{1}=\"PREMIUM\",{2}{1}=\"YES\",{3}{1}<'azure-premium-disk-prices'!B{4},{3}{1}>0),1+IF({5}{1}=\"YES\",1),\"\")".format(dataDiskTypeColumn, rowIndex, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab, ASRColumn)
		else:
			formula ="=IF(AND({0}{1}=\"PREMIUM\",{2}{1}=\"YES\",{3}{1}>'azure-premium-disk-prices'!B{4},{3}{1}<'azure-premium-disk-prices'!B{6}),1+IF({5}{1}=\"YES\",1),\"\")".format(dataDiskTypeColumn, rowIndex, dataOKColumn, dataDiskSizeColumn, diskDataIndexInDiskTab - 1, ASRColumn, diskDataIndexInDiskTab)
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
customerVMDataExcelTab.set_column(xls.managedStandardOSDiskColumn['firstColumnIndex'], xls.managedStandardOSDiskColumn['firstColumnIndex'], xls.managedStandardOSDiskColumn['width']) 
customerVMDataExcelTab.set_column(xls.managedPremiumOSDiskColumn['firstColumnIndex'],  xls.managedPremiumOSDiskColumn['firstColumnIndex'],  xls.managedPremiumOSDiskColumn['width']) 
	#SET HEADER
customerVMDataExcelTab.write(xls.managedStandardOSDiskColumn['firstCellRow'], xls.managedStandardOSDiskColumn['firstColumnIndex'], xls.managedStandardOSDiskColumn['name'] , selectHeaderStyle)
customerVMDataExcelTab.write(xls.managedPremiumOSDiskColumn['firstCellRow'],  xls.managedPremiumOSDiskColumn['firstColumnIndex'] , xls.managedPremiumOSDiskColumn['name']  , selectHeaderStyle)
	#ROWS
formulaDiskOSStandardPattern="=IF(AND({0}{1}=\"STANDARD\", {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1) ,\"\")"
formulaDiskOSPremiumPattern="=IF( AND({0}{1}=\"PREMIUM\",  {2}{1}=\"YES\"), IF({3}{1}=\"YES\",2,1) ,\"\")"

for rowIndex in range(1,xls.rowsForVMInput):	
	#COUNT OS STANDARD
	formulaDiskOSStandard=formulaDiskOSStandardPattern.format(osDiskTypeColumn , rowIndex+1, dataOKColumn, ASRColumn)
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedStandardOSDiskColumn['firstColumnIndex'], formulaDiskOSStandard, selectBodyStyle)
	#COUNT OS PREMIUM
	formulaDiskOSPremium =formulaDiskOSPremiumPattern.format(osDiskTypeColumn  , rowIndex+1, dataOKColumn, ASRColumn)
	customerVMDataExcelTab.write_formula(rowIndex, xls.managedPremiumOSDiskColumn['firstColumnIndex'] , formulaDiskOSPremium, selectBodyStyle)
	
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
	formulaPriceDisk="=B{0}*'azure-standard-disk-prices'!C{1}".format(currentCountRow, currentDiskPriceRow)
	
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 1, formulaCountDisk, selectBodyStyle)
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 2, formulaPriceDisk, selectBodyStyle)
	#PREMIUM DATA DISKS SUMMARY CALCULATIONS
for columnIndex in range(xls.managedDataDiskColumns['firstColumnIndex']  + len(priceReaderManagedDisk.standardDiskSizes), xls.managedDataDiskColumns['firstColumnIndex'] + len(priceReaderManagedDisk.standardDiskSizes) + len(priceReaderManagedDisk.premiumDiskSizes)):
	formulaCountDisk="=SUM({0}{1}:{0}{2})".format( xls.alphabet[columnIndex], xls.managedDataDiskColumns['firstCellRow'] + 1, xls.rowsForVMInput + 1)
	currentCountRow=columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] + xls.dataDiskSummary['firstCellRow'] + 1
	currentDiskPriceRow= columnIndex - xls.managedDataDiskColumns['firstColumnIndex'] - len(priceReaderManagedDisk.standardDiskSizes) + 2
	formulaPriceDisk="=B{0}*'azure-premium-disk-prices'!C{1}".format(currentCountRow, currentDiskPriceRow)
	
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 1, formulaCountDisk, selectBodyStyle)
	customerVMDataExcelTab.write_formula(currentCountRow - 1, 2, formulaPriceDisk, selectBodyStyle)

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
standardOSDiskCountFormula="=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedStandardOSDiskColumn['firstColumnIndex']], xls.managedStandardOSDiskColumn['firstCellRow'] + 2, xls.rowsForVMInput + 1)
premiumOSDiskCountFormula= "=SUM({0}{1}:{0}{2})".format( xls.alphabet[xls.managedPremiumOSDiskColumn['firstColumnIndex']] , xls.managedPremiumOSDiskColumn['firstCellRow'] + 2 , xls.rowsForVMInput + 1)
standardOSDiskPriceFormulaPattern="=B{}*'azure-standard-disk-prices'!C2"
premiumOSDiskPriceFormulaPattern= "=B{}*'azure-premium-disk-prices'!C2"
for row in xls.OSDiskSummary['rows']:
	currentRowIndex = xls.OSDiskSummary['firstCellRow'] - 1 + xls.OSDiskSummary['rows'][row]['order']
	if row == 'standard':
		countFormula=standardOSDiskCountFormula
		priceFormula=standardOSDiskPriceFormulaPattern.format(currentRowIndex + 1)
	else:
		countFormula=premiumOSDiskCountFormula	
		priceFormula=premiumOSDiskPriceFormulaPattern.format(currentRowIndex + 1)

	customerVMDataExcelTab.write(currentRowIndex, xls.OSDiskSummary['firstCellColumn'], xls.OSDiskSummary['rows'][row]['name'], selectHeaderStyle)
	customerVMDataExcelTab.write_formula(currentRowIndex, xls.OSDiskSummary['firstCellColumn'] + 1 , countFormula, selectBodyStyle)
	customerVMDataExcelTab.write_formula(currentRowIndex, xls.OSDiskSummary['firstCellColumn'] + 2 , priceFormula, selectBodyStyle)
	
#11 - BLOCK 9 - COST SUMMARY
formulaTotalComputeCost="=SUM({0}1:{0}{1})".format(columnBestYearVMPrice, xls.rowsForVMInput + 1)
formulaTotalDiskCost="=12*( SUM(C{0}:C{1}) + SUM({4}{2}:{4}{3})*'azure-standard-disk-prices'!C2 + SUM({5}{2}:{5}{3})*'azure-premium-disk-prices'!C2)".format(xls.dataDiskSummary['firstCellRow'] + 1 , xls.dataDiskSummary['firstCellRow'] + 1 +  totalNumDisks - 1, xls.managedDataDiskColumns['firstCellRow'] + 1 , xls.rowsForVMInput + 1, xls.alphabet[xls.managedStandardOSDiskColumn['firstColumnIndex']], xls.alphabet[xls.managedPremiumOSDiskColumn['firstColumnIndex']])
formulaTotalASRCost ="=12*SUM({0}{1}:{0}{2})*'azure-asr-prices'!A2".format(xls.getColumnLetterFromIndex(xls.ASRColumns['firstColumnIndex']), xls.ASRColumns['firstCellRow'] + 1, xls.rowsForVMInput + 1)
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
	formulaCheapestPricePattern="={0}{1}"
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



####################################################################
############################ OTHER TABS	############################
####################################################################
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
	auxFormula='=CONCATENATE(C{0},E{0},D{0})'.format(currentLineBase+1)
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
		auxFormula='=CONCATENATE(C{0},E{0},D{0})'.format(currentLine1Y+1)
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
		auxFormula='=CONCATENATE(C{0},E{0},D{0})'.format(currentLine3Y+1)
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