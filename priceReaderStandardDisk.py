import urllib.request
import json

urlPriceManagedDiskPublicAPI='https://azure.microsoft.com/api/v2/pricing/managed-disks/calculator/?culture=en-us&discount=mosp'
region = 'europe-west'

def getPriceMatrix():
	with urllib.request.urlopen(urlPriceManagedDiskPublicAPI) as url:
		dataBasePrice = json.loads(url.read().decode())
		regionSizes  = {data['size']:{'price':data['prices'][region]['value'], 'name':diskName} for (diskName,data) in dataBasePrice['offers'].items() if 'standard-s' in diskName and 'standard-snapshot' not in diskName and region in data['prices']}

	return regionSizes
	