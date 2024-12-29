# Tenable-Automation-Script
### This repository is created to upload all the scripts that I created in terms of Tenable/Nessus export to automate the repetitive work that I have been doing.

## IP_Exporter.py
This script will filter IP matches from columns **IP address** and copy the complete rows to a new file from a txt file with the sheet name same as the txt file name. If an IP is not found then it will create a new CSV file with the list of IPs that were not present in the master sheet

Requirements for the Script to work
```
* Minimum one txt file with the list of IPs or multiple will also work
```

Below is the flow and the script,
* Will expect two inputs from the user,
  * Master file from which script will filter the IPs to copy to a new file
  * The filename of the output to be stored
* Will look for a list of txt files and reads the list of IPs.
* A new xlsx file will be created based on the user input
* With the txt file name as a sheet name all the rows that matches the IPs available in the master sheet will be exported to a new file
* If there are multiple txt files, then the script will create multiple sheets in the workbook based on the txt filename.



## CSV_EXCEL_combiner.py
This script is created to combine all my list of CSV's into one excel file
Requirements for the Script to work
```
* Minimum two CSV files to combine
Note: Excel has a limitation of cell value maximum to 32000 lines whereas CSV doesn't have when opened using a text editor
```

Below is the flow and the script,
* Will look for a list of CSVs in a folder and create a new xlsx file combining all the csv's
