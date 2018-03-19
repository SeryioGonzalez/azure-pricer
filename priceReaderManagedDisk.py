import urllib.request
import json

urlPriceManagedDiskPublicAPI='https://azure.microsoft.com/api/v2/pricing/managed-disks/calculator/?culture=en-us&discount=mosp&currency=eur'

premiumDiskSizes  = ['P4', 'P6', 'P10', 'P15', 'P20', 'P30', 'P40', 'P50']
standardDiskSizes = ['S4', 'S6', 'S10', 'S15', 'S20', 'S30', 'S40', 'S50']

#{size: { price:N, name:name }}
def getPriceMatrixStandard(region):
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		regionSizes  = {data['size']:{'price':data['prices'][region]['value'], 'name':diskName} for (diskName,data) in dataBasePrice['offers'].items() if 'standard-s' in diskName and 'standard-snapshot' not in diskName and region in data['prices']}

	return regionSizes
	
#{size: { price:N, name:name }}
def getPriceMatrixPremium(region):
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		regionSizes  = {data['size']:{'price':data['prices'][region]['value'], 'name':diskName} for (diskName,data) in dataBasePrice['offers'].items() if 'premium-p' in diskName and region in data['prices']}

	return regionSizes
