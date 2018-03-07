alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ' ]

customerInputColumns = {
	'VM NAME' : {
		'alias' : 'VM NAME'
		'width' : 20,
		'position' : 0
	},
	'CPUs' : {
		'alias' : 'CPUs',
		'width' : 5,
		'position' : 1
	},
	'Mem(GB)' : {
		'alias' : 'Mem(GB)',
		'width' : 9,
		'position' : 2
	},
	'DATA STORAGE' : {
		'alias' : 'DATA STORAGE',
		'width' : 14,
		'position' : 3
	},
	'DATA STORAGE TYPE' : {
		'alias' : 'DATA STORAGE TYPE',
		'width' : 19,
		'position' : 4,
		'default' : 'STANDARD'
	},
	'OS STORAGE TYPE' : {
		'alias' : 'OS STORAGE TYPE',
		'width' : 17,
		'position' : 5,
		'default' : 'STANDARD'
	},
	'SAP' : {
		'alias' : 'SAP',
		'width' : 5,
		'position' : 6,
		'default' : 'NO'
	},
	'GPU' : {
		'alias' : 'GPU',
		'width' : 5,
		'position' : 7
		'default' : 'NO'
	},
	'ASR' : {
		'alias' : 'ASR',
		'width' : 5,
		'position' : 8
		'default' : 'NO'
	},
	'HOURS/MONTH' : {
		'alias': 'HOURS/MONTH',
		'width' : 15,
		'position' : 9
		'default' : '730'
	},
	'USE B SERIES' : {
		'alias' : 'USE B SERIES',
		'width' : 12,
		'position' : 10
		'default' : 'NO'
	},
	'ALL DATA OK' : {
		'alias' : 'ALL DATA OK',
		'width' : 12,
		'position' : 11
	}
}
