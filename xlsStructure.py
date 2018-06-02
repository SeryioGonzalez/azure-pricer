import priceReaderManagedDisk 

class xlsStructure:

	rowsForVMInput=1000

	alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL' ]
	
	firstColumnWidth=10
	firstColumnIndex=0
	
	regionListColumn=63
	
	#BLOCK 1
	assumptions = {
		'firstCellColumn':0,
		'firstCellRow':0,		
		'header': {
			'width': 2,
			'title': 'ASSUMPTIONS'
		},
		'rows': {
			'PERF':    {'name': 'PERF GAIN', 'order':1, 'default':0},
			'MODE':    {'name': 'MODE'     , 'order':2, 'default':'CPU+MEM', 'validationList' : ['CPU+MEM', 'MEM']}
		}
	}
	
	#BLOCK 2
	customerInputColumns = {
		'firstCellRow':0,
		'firstColumnIndex' : 4,	
		'columns': {
			'VM NAME' : {
				'alias' : 'VM NAME',
				'width' : 20,
				'index' : 0,
				'default' : ''
			},
			'CPUs' : {
				'alias' : 'CPUs',
				'width' : 5,
				'index' : 1,
				'default' : ''
			},
			'Mem(GB)' : {
				'alias' : 'Mem(GB)',
				'width' : 9,
				'index' : 2,
				'default' : ''
			},
			'DATA STORAGE' : {
				'alias' : 'DATA STORAGE(GB)',
				'width' : 19,
				'index' : 3,
				'default' : ''
			},
			'REGION' : {
				'alias' : 'REGION',
				'width' : 20,
				'index' : 4,
				'default' : 'europe-west',
				'validationList' : '=$' + alphabet[regionListColumn] + '$1:$' + alphabet[regionListColumn] + '$36'
			},
			'LICENSE' : {
				'alias' : 'LICENSE',
				'width' : 9,
				'index' : 5,
				'default' : 'linux',
				'validationList' : ['linux', 'windows']
			},
			'DATA STORAGE TYPE' : {
				'alias' : 'DATA STORAGE TYPE',
				'width' : 18,
				'index' : 6,
				'default' : 'STANDARD',
				'validationList' : ['STANDARD', 'PREMIUM']
			},
			'OS DISK' : {
				'alias' : 'OS DISK',
				'width' : 8,
				'index' : 7,
				'default' : 'S4',
				'validationList' : ['S4', 'S6', 'S10', 'P4', 'P6', 'P10']
			},
			'SAP' : {
				'alias' : 'SAP',
				'width' : 4,
				'index' : 8,
				'default' : 'NO',
				'validationList' : ['YES', 'NO']
			},
			'GPU' : {
				'alias' : 'GPU',
				'width' : 4,
				'index' : 9,
				'default' : 'NO',
				'validationList' : ['YES', 'NO']
			},
			'ASR' : {
				'alias' : 'ASR',
				'width' : 4,
				'index' : 10,
				'default' : 'NO',
				'validationList' : ['YES', 'NO']
			},
			'HOURS/MONTH' : {
				'alias': 'HOURS/MONTH',
				'width' : 15,
				'index' : 11,
				'default' : 730
			},
			'USE B SERIES' : {
				'alias' : 'USE B SERIES',
				'width' : 12,
				'index' : 12,
				'default' : 'NO',
				'validationList' : ['YES', 'NO']
			},
			'RESERVED INST' : {
				'alias' : 'RESERVED INST.',
				'width' : 15,
				'index' : 13,
				'default' : 'YES ALL',
				'validationList' : ['YES ALL', 'YES 1Y', 'NO']
			},
			'ALL DATA OK' : {
				'alias' : 'DATA OK',
				'width' : 8,
				'index' : 14,
				'default' : ''
			}
		}	
	}
	#BLOCK 10
	computeSummaryColumns = {
		'firstCellRow':0,
		'firstColumnIndex' : customerInputColumns['firstColumnIndex'] + len(customerInputColumns['columns']),	
		'columns': {
			'CHEAPEST SIZE' : {
				'alias' : 'CHEAPEST VM SIZE',
				'width' : 17,
				'index' : 0
			},
			'CHEAPEST PRICE' : {
				'alias' : 'CHEAPEST VM PRICE',
				'width' : 18,
				'index' : 1
			},
			'CHEAPEST MODEL' : {
				'alias' : 'CHEAPEST MODEL',
				'width' : 16,
				'index' : 2
			}
		}	
	}
	#BLOCK 3	
	VMCalculationColumns = {
		'firstCellRow':0,
		'firstColumnIndex' : computeSummaryColumns['firstColumnIndex'] + len(computeSummaryColumns['columns']),			
		'columns': {
			'BEST SIZE PAYG' : {
				'alias' : 'BEST SIZE PAYG',
				'width' : 14,
				'index' : 0
			},
			'PRICE(H) PAYG' : {
				'alias' : 'PRICE(H) PAYG',
				'width' : 14,
				'index' : 1
			},
			'BEST SIZE 1Y' : {
				'alias' : 'BEST SIZE 1Y',
				'width' : 14,
				'index' : 2
			},
			'PRICE(H) 1Y' : {
				'alias' : 'PRICE(H) 1Y',
				'width' : 11,
				'index' : 3
			},
			'BEST SIZE 3Y' : {
				'alias' : 'BEST SIZE 3Y',
				'width' : 14,
				'index' : 4
			},
			'PRICE(H) 3Y' : {
				'alias' : 'PRICE(H) 3Y',
				'width' : 11,
				'index' : 5
			},
			'PAYG' : {
				'alias' : 'PAYG 1Y',
				'width' : 11,
				'index' : 6
			},
			'1Y RI' : {
				'alias' : '1Y RI',
				'width' : 10,
				'index' : 7
			},
			'3Y RI' : {
				'alias' : '3Y RI',
				'width' : 10,
				'index' : 8
			},
			'BEST PRICE' : {
				'alias': 'BEST PRICE',
				'width' : 10,
				'index' : 9
			}
		}	
	}

	#BLOCK 4		
	managedDataDiskColumns = {
		'firstCellRow': 0,	
		'prefix' : 'DATA - ',
		'width' : 10,
		'firstColumnIndex' : VMCalculationColumns['firstColumnIndex'] + len(VMCalculationColumns['columns'])
	}	
	
	#BLOCK 5
	ASRColumns = {
		'firstCellRow': 0,	
		'name' : 'ASR',
		'width' : 8,
		'firstColumnIndex' : managedDataDiskColumns['firstColumnIndex'] + len(priceReaderManagedDisk.standardDiskSizes )  + len(priceReaderManagedDisk.premiumDiskSizes )
	}	
	
	#BLOCK 6
	managedS4OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK S4',
		'width' : 11,
		'firstColumnIndex' : ASRColumns['firstColumnIndex'] + 1
	}
	managedP4OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK P4',
		'width' : 11,
		'firstColumnIndex' : managedS4OSDiskColumn['firstColumnIndex'] + 1
	}
	
	managedS6OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK S6',
		'width' : 11,
		'firstColumnIndex' : managedP4OSDiskColumn['firstColumnIndex'] + 1
	}
	
	managedP6OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK P6',
		'width' : 11,
		'firstColumnIndex' : managedS6OSDiskColumn['firstColumnIndex'] + 1
	}
	
	managedS10OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK S10',
		'width' : 11,
		'firstColumnIndex' : managedP6OSDiskColumn['firstColumnIndex'] + 1
	}
	
	managedP10OSDiskColumn = {
		'firstCellRow': 0,	
		'name' : 'OS DISK P10',
		'width' : 11,
		'firstColumnIndex' : managedS10OSDiskColumn['firstColumnIndex'] + 1
	}
	
	ssdRequiredCheckColumn = {
		'firstCellRow': 0,	
		'name' : 'SSD NEEDED',
		'width' : 11,
		'firstColumnIndex' : managedP10OSDiskColumn['firstColumnIndex'] + 1
	}
	
	asrPriceColumn = {
		'firstCellRow': 0,	
		'name' : 'ASR PRICE',
		'width' : 10,
		'firstColumnIndex' : ssdRequiredCheckColumn['firstColumnIndex'] + 1
	}

	diskPriceColumn = {
		'firstCellRow': 0,	
		'name' : 'DISK PRICE',
		'width' : 12,
		'firstColumnIndex' : asrPriceColumn['firstColumnIndex'] + 1
	}
	
	#BLOCK 7
	dataDiskSummary = {
		'firstCellColumn':0,
		'firstCellRow': 12,
		'header': {
			'width': 2,
			'title': 'DATA DISK SUMMARY'
		},
		'columns': {
			'DISK SIZE': {'name': 'DISK SIZE', 'order':1},
			'COUNT':     {'name': 'COUNT'    , 'order':2}
		}
	}

	#BLOCK 8
	OSDiskSummary = {
		'firstCellColumn':0,
		'firstCellRow': 31,
		'header': {
			'width': 2,
			'title': 'OS DISK SUMMARY'
		},
		'columns': {
			'DISK SIZE': {'name': 'DISK SIZE', 'order':1},
			'COUNT':     {'name': 'COUNT'    , 'order':2}
		},
		'rows': {
			'S4':  {'name': 'S4', 'order':1},
			'P4':  {'name': 'P4', 'order':2},
			'S6':  {'name': 'S4', 'order':3},
			'P6':  {'name': 'P4', 'order':4},
			'S10': {'name': 'S10', 'order':5},
			'P10': {'name': 'P10', 'order':6}		
		}
	}
	
	#BLOCK 9	
	costSummary = {
		'firstCellColumn':0,
		'firstCellRow':4,		
		'header': {
			'width': 2,
			'title': 'YEAR TOTALS(€)'
		},
		'rows': {
			'COMPUTE': {'name': 'COMPUTE', 'order':1},
			'STORAGE': {'name': 'STORAGE', 'order':2},
			'ASR':     {'name': 'ASR',     'order':3},
			'TOTAL':   {'name': 'TOTAL',   'order':4}
		}
	}
	
	premiumDisks = [
		{ "diskName":"P4",  "diskSize":32},
		{ "diskName":"P6",  "diskSize":64},
		{ "diskName":"P10", "diskSize":128},
		{ "diskName":"P15", "diskSize":256},
		{ "diskName":"P20", "diskSize":512},
		{ "diskName":"P30", "diskSize":1024},
		{ "diskName":"P40", "diskSize":2048},
		{ "diskName":"P50", "diskSize":4096}		
	]
	
	standardDisks = [
		{ "diskName":"S4",  "diskSize":32},
		{ "diskName":"S6",  "diskSize":64},
		{ "diskName":"S10", "diskSize":128},
		{ "diskName":"S15", "diskSize":256},
		{ "diskName":"S20", "diskSize":512},
		{ "diskName":"S30", "diskSize":1024},
		{ "diskName":"S40", "diskSize":2048},
		{ "diskName":"S50", "diskSize":4096}		
	]
	
	
	#GIVEN A CUSTOMER DATA COLUMN RELATIVE INDEX, GET ABSOLUTE SPREADSHEET POSITION
	def getCustomerDataColumnPositionInExcel(columnIndex):
		return columnIndex + xlsStructure.customerInputColumns['firstColumnIndex']

	#GIVEN A CALCULATION COLUMN RELATIVE INDEX, GET ABSOLUTE SPREADSHEET POSITION		
	def getCalculationColumnPositionInExcel(columnIndex):
		return columnIndex + xlsStructure.VMCalculationColumns['firstColumnIndex']

	#GIVEN AN ABSOLUTE INTEGER INDEX, GET COLUMN
	def getColumnLetterFromIndex(columnIndex):
		return xlsStructure.alphabet[columnIndex]	

	#GIVEN AN ASSUMPTION NAME, GET ITS VALUE CELL		
	def getAssumptionValueCell(assumption, fixed=True):
		firstAssumptionRow = xlsStructure.assumptions['firstCellRow']
		firstAssumptionColumn = xlsStructure.assumptions['firstCellColumn']
		relativeAssumptionRow = xlsStructure.assumptions['rows'][assumption]['order']
		
		assumptionRow = firstAssumptionRow + relativeAssumptionRow
		assumptionColumn = firstAssumptionColumn + 1
		assumptionColumnLetter = xlsStructure.alphabet[assumptionColumn]
		if (fixed):
			cellInText = "$" + assumptionColumnLetter + "$" + str(assumptionRow + 1)
		else:
			cellInText = assumptionColumnLetter + str(assumptionRow + 1)
		
		return cellInText
	
	def getVMCalculationColumn(columnName):
		categoryIndexInVMCalculations = xlsStructure.VMCalculationColumns['columns'][columnName]['index']
		columnIndex = xlsStructure.getCalculationColumnPositionInExcel(categoryIndexInVMCalculations)
		columnLetter = xlsStructure.getColumnLetterFromIndex(columnIndex)
		
		return columnLetter
		
		
