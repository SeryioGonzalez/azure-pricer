# Azure Quote Estimation Tool

Unofficial tool for providing a first order estimation on Azure migration projects to Western Europe.
It generates a spreadsheet for performing calculations

For creating a new spreadsheet:<br/>
```
$ python3 xlsGenerator.py [nameOfOutputXLSFile]
```

The tool will generate a new spreadsheet, equivalent to <a href="https://github.com/SeryioGonzalez/azure-pricer/raw/master/aux/Azure-Quote-Tool-020418.xlsx">Azure-Quote-Tool-020418.xlsx</a>

# Considerations

- This is a unofficial tool for providing a first order estimation on Azure migration projects

- It contains Azure pricing information read from Azure pricing APIs at the day of the spreadsheet creation

- All prices are for Western Europe deployments in this version

- The tool consider IaaS projects and estimates compute, storage and ASR costs

# Services considered for a quote

- Managed Disks for OS and data drives
- Premium or Standard Disk for OS and data dries
- ASR Costs per VM
- ASR Disk duplication costs
- VMs with no software license, i.e., quoting Ubuntu VMs
- PAYG(hour), 3 or 1 year Reserved Instance pricing models
- VMs certified for SAP NetWeaver loads
- VMs equipped with GPU cards
- Usage of burstable VMs
- Percentual performance improvement in Azure
	- If I input 20% performance win per CPU and memory on Azure, a VM request for 100 GB RAM will look VMs sises starting at 80 GB

# Additional points to be added for a full quote

- VPN or Express Route (ER) Gateway costs
- If ER is used, the cost of the ER circuit in Azure and telco costs
- Outgoing bandwidth costs
- VM Backups
- Cost for services other than compute, managed disks and ASR
