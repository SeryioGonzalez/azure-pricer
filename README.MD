# Azure Quote Estimation Tool

The tool generates a spreadsheet like <a href="https://github.com/SeryioGonzalez/azure-pricer/raw/master/output/Azure-Quote-Tool-20220826.xlsx">Azure-Quote-Tool-20220826.xlsx</a>

Unofficial tool for an estimation on Azure VM migration projects.
It generates a spreadsheet for performing calculations

For creating a new spreadsheet:<br/>
```
#ONLY IF YOU HAD NOT INSTALLED PIP3 AND THE XLSWRITER LIBRARY
$ sudo apt install -y python3-pip
$ pip3 install xlsxwriter

#COMMAND REQUIRED FOR CREATING A SPREADSHEET
# python3 xls_generator.py [nameOfOutputXLSFile]
$ python3 xls_generator.py Azure-Quote-Tool-$(date +%d%m%y).xlsx
$ ls -1
Azure-Quote-Tool-020618.xlsx
```

It works for Excel 2016 and beyond

# Considerations

- This is a unofficial tool for providing a first order estimation on Azure migration projects

- It contains Azure pricing information read from Azure pricing APIs at the day of the spreadsheet creation

- The tool consider IaaS projects and estimates compute, storage and ASR costs

# Services considered for a quote

- Managed Disks for OS and data drives
- Premium or Standard Managed Disk for OS and data drives
	- OS Disk sizes limited to P4, P4, P10, S4, S6 or S10
- ASR Costs per VM
- ASR Disk duplication costs
- PAYG(hour), 3 or 1 year Reserved Instance pricing models
- VMs certified for SAP NetWeaver loads
- VMs equipped with GPU cards
- Usage of burstable VMs
- Percent performance improvement in Azure
	- If I input 20% performance win per CPU and memory on Azure, a VM request for 100 GB RAM will look VMs sizes starting at 80 GB
- Selection based on CPU and Memory or just Memory

# Additional points to be added for a full quote

- VPN or Express Route (ER) Gateway costs
- If ER is used, the cost of the ER circuit in Azure and telco costs
- Outgoing bandwidth costs
- VM Backups
- Cost for services other than compute, managed disks and ASR
