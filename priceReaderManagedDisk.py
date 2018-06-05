import urllib.request
import json

urlPriceManagedDiskPublicAPI='https://azure.microsoft.com/api/v2/pricing/managed-disks/calculator/?culture=en-us&discount=mosp&currency=eur'

premiumDiskSizes  = ['P4', 'P6', 'P10', 'P15', 'P20', 'P30', 'P40', 'P50']
standardDiskSizes = ['S4', 'S6', 'S10', 'S15', 'S20', 'S30', 'S40', 'S50']

def getPriceMatrixStandard(regions):
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		
	allRegionsSizes = {}	
	keyWord='standardhdd-s'
	
	for region in regions:
		regionSizes  = {region + "-" + str(data['size']) : {'price':data['prices'][region]['value'], 'name':diskName.upper().split("-")[1], 'region':region, 'size':data['size']} for (diskName,data) in dataBasePrice['offers'].items() if keyWord in diskName and 'snapshot' not in diskName and region in data['prices']}
		allRegionsSizes.update(regionSizes)

	return allRegionsSizes

def getPriceMatrixPremium(regions):
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		
	allRegionsSizes = {}	
	keyWord='premiumssd-p'
	
	for region in regions:
		regionSizes  = {region + "-" + str(data['size']) : {'price':data['prices'][region]['value'], 'name':diskName.upper().split("-")[1], 'region':region, 'size':data['size']} for (diskName,data) in dataBasePrice['offers'].items() if keyWord in diskName and region in data['prices']}
		allRegionsSizes.update(regionSizes)

	return allRegionsSizes
