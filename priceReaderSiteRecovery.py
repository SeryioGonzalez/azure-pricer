import urllib.request
import json

urlPriceManagedDiskPublicAPI='https://azure.microsoft.com/api/v2/pricing/site-recovery/calculator/?culture=en-us&discount=mosp&currency=eur'

#{region: price}

def getPriceMatrix(regions):
	
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		
	allRegionsSizes = {}	

	for region in regions:
		thisRegionSize  = {region:price['prices'][region]['value'] for (siteRecoveryName,price) in dataBasePrice['offers'].items() if 'recover-to-azure' in siteRecoveryName and region in price['prices']}
		allRegionsSizes.update(thisRegionSize)
	
	return allRegionsSizes