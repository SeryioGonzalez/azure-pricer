
class xlsStructure:

	firstColumnCustomerInput=3
	rowsForVMInput=250

	alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ' ]
		
	assumptions = {
		'firstCellColumn':0,
		'firstCellRow':1,		
		'header': {
			'width': 2,
			'title': 'ASSUMPTIONS'
		},
		'rows': {
			'PERF':    {'name': 'PERF. TOLERANCE', 'order':1},
			'RESINST': {'name': 'RES. INST.',      'order':2},
			'USD2EURO':{'name': 'USD TO EURO',     'order':3}
		}
	}
	
	costSummary = {
		'firstCellColumn':0,
		'firstCellRow':7,		
		'header': {
			'width': 2,
			'title': 'YEAR TOTALS (EUROS)'
		},
		'rows': {
			'COMPUTE': {'name': 'COMPUTE', 'order':1},
			'STORAGE': {'name': 'STORAGE', 'order':2},
			'ASR':     {'name': 'ASR',     'order':3},
			'TOTAL':   {'name': 'TOTAL',   'order':4}
		}
	}
	
	customerInputColumns = {
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
			'alias' : 'DATA STORAGE',
			'width' : 14,
			'index' : 3,
			'default' : ''
		},
		'DATA STORAGE TYPE' : {
			'alias' : 'DATA STORAGE TYPE',
			'width' : 18,
			'index' : 4,
			'default' : 'STANDARD',
			'validationList' : ['STANDARD', 'PREMIUM']
		},
		'OS STORAGE TYPE' : {
			'alias' : 'OS STORAGE TYPE',
			'width' : 16,
			'index' : 5,
			'default' : 'STANDARD',
			'validationList' : ['STANDARD', 'PREMIUM']
		},
		'SAP' : {
			'alias' : 'SAP',
			'width' : 4,
			'index' : 6,
			'default' : 'NO',
			'validationList' : ['YES', 'NO']
		},
		'GPU' : {
			'alias' : 'GPU',
			'width' : 4,
			'index' : 7,
			'default' : 'NO',
			'validationList' : ['YES', 'NO']
		},
		'ASR' : {
			'alias' : 'ASR',
			'width' : 4,
			'index' : 8,
			'default' : 'NO',
			'validationList' : ['YES', 'NO']
		},
		'HOURS/MONTH' : {
			'alias': 'HOURS/MONTH',
			'width' : 15,
			'index' : 9,
			'default' : 730
		},
		'USE B SERIES' : {
			'alias' : 'USE B SERIES',
			'width' : 12,
			'index' : 10,
			'default' : 'NO',
			'validationList' : ['YES', 'NO']
		},
		'ALL DATA OK' : {
			'alias' : 'DATA OK',
			'width' : 8,
			'index' : 11,
			'default' : ''
		}
	}

	VMCalculationColumns = {
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
			'width' : 11,
			'index' : 2
		},
		'PRICE(H) 1Y' : {
			'alias' : 'PRICE(H) 1Y',
			'width' : 11,
			'index' : 3
		},
		'BEST SIZE 3Y' : {
			'alias' : 'BEST SIZE 3Y',
			'width' : 11,
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
	
	managedDataDiskColumns = {
		'prefix' : 'DATA - ',
		'width' : 10,
		'firstColumnIndex' : firstColumnCustomerInput + len(customerInputColumns) + len(VMCalculationColumns)
	}
	
	
	def getCustomerDataColumnPositionInExcel(columnIndex):
		return columnIndex + xlsStructure.firstColumnCustomerInput
		
	def getCalculationColumnPositionInExcel(columnIndex):
		return columnIndex + xlsStructure.firstColumnCustomerInput + len(xlsStructure.customerInputColumns)

	def getColumnLetterFromIndex(columnIndex):
		return alphabet[columnIndex]	
	
	def getAssumptionValueCell(assumption):
		firstAssumptionRow = xlsStructure.assumptions['firstCellRow']
		firstAssumptionColumn = xlsStructure.assumptions['firstCellColumn']
		relativeAssumptionRow = xlsStructure.assumptions['rows'][assumption]['order']
		
		assumptionRow = firstAssumptionRow + relativeAssumptionRow
		assumptionColumn = firstAssumptionColumn + 1
		assumptionColumnLetter = xlsStructure.alphabet[assumptionColumn]
		
		cellInText = "$" + assumptionColumnLetter + "$" + str(assumptionRow + 1)

		return cellInText
		