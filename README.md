# PS_Printer_Ink_Level_SNMP
Create a HTML report of printer toner level using SNMP and PowerShell

This script will generate an html page containing all printers described in the printerlist.txt file
For each printer, it will get the level for all cartridge + any existing alarm

It will retrieve information using SNMP and the default printer MIB
  Input:
       printerlist.txt : csv file containing printers where information will be retrieved.
                         each line contains value,name,description
                         with 
                             value: DNS name of your printer
                             name: Name of the printer as displayed in the report
                             description: Description of the printer that will appear in the report
