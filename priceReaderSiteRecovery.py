import urllib.request
import json

urlPriceManagedDiskPublicAPI='https://azure.microsoft.com/api/v2/pricing/site-recovery/calculator/?culture=en-us&discount=mosp'

#{recoveryName: price}

def getPriceMatrix(region):
	
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		regionSizes  = {siteRecoveryName:price['prices'][region]['value'] for (siteRecoveryName,price) in dataBasePrice['offers'].items() if 'recover-to-azure' in siteRecoveryName and region in price['prices']}

	return regionSizes
