import urllib.request
import json

sapVMs=['a5', 'a6', 'a7', 'a8', 'a9', 'a10', 'a11', 'd11', 'd12', 'd13', 'd14', 'ds11', 'ds12', 'ds13', 'ds14', 'ds11v2', 'ds12v2', 'ds13v2', 'ds14v2', 'ds15v2', 'gs1', 'gs2', 'gs3', 'gs4', 'gs5', 'm64ms', 'm128ms', 'm64s', 'm128s']

urlPriceBasePublicAPI='https://azure.microsoft.com/api/v2/pricing/virtual-machines-base/calculator/?culture=en-us&discount=mosp'
urlPrice1YeaPublicAPI='https://azure.microsoft.com/api/v2/pricing/virtual-machines-base-one-year/calculator/?culture=en-us&discount=mosp'
urlPrice3YeaPublicAPI='https://azure.microsoft.com/api/v2/pricing/virtual-machines-base-three-year/calculator/?culture=en-us&discount=mosp'

region = 'europe-west'

def get3YeaPrice(sizeName, regionSizes3Year):
	try:
		price = regionSizes3Year[sizeName]
	except:
		price = 1000000
	return price
	
def get1YeaPrice(sizeName, regionSizes1Year):
	try:
		price = regionSizes1Year[sizeName]
	except:
		price = 1000000
	return price

def getSapCapable(sizeName):
	if sizeName in sapVMs:
		return 'YES'
	else:
		return 'NO'
	
def getGPUCapable(sizeData):
	try:
		sizeData['gpu']
		gpus = 'YES'
	except:
		gpus = 'NO'
		
	return gpus
	
def cleanSizeName(sizeName):
	try:
		newName = sizeName.split('-')[1]
	except:
		newName = sizeName
	return newName

def flagBurstable(formattedVMSize):
	if formattedVMSize[0] == 'b':
		return "YES"
	else:
		return "NO"
	
def getPriceMatrix():
	
	with urllib.request.urlopen(urlPrice1YeaPublicAPI) as url:
		data1YeaPrice = json.loads(url.read().decode())
		regionSizes1Year = {k:v['prices'][region]['value'] for (k,v) in data1YeaPrice['offers'].items() if 'linux' in k and 'standard' in k and region in v['prices']}	
		
	with urllib.request.urlopen(urlPrice3YeaPublicAPI) as url:
		data3YeaPrice = json.loads(url.read().decode())
		regionSizes3Year = {k:v['prices'][region]['value'] for (k,v) in data3YeaPrice['offers'].items() if 'linux' in k and 'standard' in k and region in v['prices']}	
	
	with urllib.request.urlopen(urlPriceBasePublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		regionSizes  = {cleanSizeName(sizeName): { \
								'payg': v['prices'][region]['value'], \
								'1y': get1YeaPrice(sizeName, regionSizes1Year), \
								'3y': get3YeaPrice(sizeName, regionSizes3Year), \
								'cpu':v['cores'], \
								'ram':v['ram'], \
								'sap': getSapCapable(cleanSizeName(sizeName)), \
								'gpu': getGPUCapable(v),
								'burstable': flagBurstable(cleanSizeName(sizeName)) 
								} for (sizeName,v) in dataBasePrice['offers'].items() if 'linux' in sizeName and 'standard' in sizeName and region in v['prices']}

	return regionSizes
