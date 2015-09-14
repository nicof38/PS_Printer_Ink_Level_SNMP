# Generate an html page containing all EMF described in the printerlist.txt file
# For each EMF, it will get the level for all cartridge + any existing alarm
# It will retrieve information using SNMP and the default printer MIB
#
#  Input:
#       printerlist.txt : csv file containing EMF where information will be retrieved.

$printerlist = import-csv .\printerlist.txt -header Value,Name,Description
$outfile_final = "\\soitec.net\files\Corporate\Reporting\Publication\IT\Support\PrinterReport.html"
$outfile = $outfile_final + ".tmp"
$outfile_back = $outfile_final + ".bak.html"
Write-Host $outfile_temp
$SNMP = new-object -ComObject olePrn.OleSNMP
$ErrorActionPreference = "Continue"
$total = ($printerlist.value|? {$_ -notlike "-*"}).count
$a = Get-Date -format "dd/MM HH:mm"

$ScriptPath = Split-Path $MyInvocation.InvocationName
$SoapMonitor = "$ScriptPath\SOAP_Monitor.ps1"

Copy-Item -Path $outfile_final -Destination $outfile_back -Force

# Write header of the HTML file 
Write "`
<html>`
<head>`
<title>Printer Report</title>`
<style>* {font-family:'Trebuchet MS';} table,th,td {border: 1px solid black;border-collapse: collapse;}</style>`
</head>`
<body>"|out-file -Encoding "UTF8" $outfile
 
write "Reporting on $total printers"
$x = 0
write "<TABLE style='width:100%'>"|add-content $outfile 
write  "<tr><th>Description</th><th>Name</th><th>Type</th><th>Black</th><th>Cyan</th><th>Magenta</th><th>Yellow</th><th>Status</th></tr>"|add-content $outfile 
foreach ($p in $printerlist){
 
    # Treat any line starting with "-" as a comment
    if ($p.value -like "-*"){
        write "<h3>",$p.value.replace('-',''),"</h3>"|add-content $outfile
        }
 
    if ($p.value -notlike "-*"){
        write "<tr>"|add-content $outfile
        $x = $x + 1
        $printertype = $nul
        $status = $nul
        $percentremaining = $nul
        $blackpercentremaining = $nul
        $cyanpercentremaining = $nul
        $magentapercentremaining = $nul
        $yellowpercentremaining = $nul
        $wastepercentremaining = $nul
 
        if (!(test-connection $p.Value -Quiet -count 1)){
            write "<td>"|add-content $outfile
            write ("<b>" + $p.description + "</b>")|add-content $outfile
            write "</td>"|add-content $outfile
            write "<td>"|add-content $outfile
            write ("<a style='text-decoration:none;font-weight:bold;' href=http://" + $p.value + " target='_new'> " + $p.value + "</a>")|add-content $outfile
            write "</td>"|add-content $outfile
            write "<td>"|add-content $outfile
            write ("<b>" + "Offline" + "</b>")|add-content $outfile            
            write "</td>"|add-content $outfile
            write "<td></td><td></td><td></td><td></td><td></td>"|add-content $outfile
            
            }

        if (test-connection $p.value -quiet -count 1){
            $snmp.open($p.value,"public",2,3000)
            $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
            write ([string]$x + ": " + [string]$p.Value + " " + $printertype)
        }
        $managed = $false
 
        if (($printertype -like "*Canon iR*") ){
            $managed = $true
	        $blacktonervolume = $snmp.get("43.11.1.1.8.1.1")
	        $blackcurrentvolume = $snmp.get("43.11.1.1.9.1.1")
	        [int]$blackpercentremaining = ($blackcurrentvolume / $blacktonervolume) * 100
	        $cyantonervolume = $snmp.get("43.11.1.1.8.1.2")
	        $cyancurrentvolume = $snmp.get("43.11.1.1.9.1.2")
	        [int]$cyanpercentremaining = ($cyancurrentvolume / $cyantonervolume) * 100
	        $magentatonervolume = $snmp.get("43.11.1.1.8.1.3")
	        $magentacurrentvolume = $snmp.get("43.11.1.1.9.1.3")
	        [int]$magentapercentremaining = ($magentacurrentvolume / $magentatonervolume) * 100
	        $yellowtonervolume = $snmp.get("43.11.1.1.8.1.4")
	        $yellowcurrentvolume = $snmp.get("43.11.1.1.9.1.4")
	        [int]$yellowpercentremaining = ($yellowcurrentvolume / $yellowtonervolume) * 100
	        $statustree = $snmp.gettree("43.18.1.1.8")
	        $status = $statustree|? {$_ -notlike "print*"} #status, including low ink warnings
	        $status = $status|? {$_ -notlike "*bypass*"}

	        $name = $snmp.get(".1.3.6.1.2.1.1.5.0")
	        if ($name -notlike "PX*"){$name = $p.name}
        }
        elseif (($printertype -like "*Officejet 6500 E710*")) {
            $managed = $true
	        $blacktonervolume = $snmp.get("43.11.1.1.8.1.1")
	        $blackcurrentvolume = $snmp.get("43.11.1.1.9.1.1")
	        [int]$blackpercentremaining = ($blackcurrentvolume / $blacktonervolume) * 100
	        $yellowtonervolume = $snmp.get("43.11.1.1.8.1.2")
	        $yellowcurrentvolume = $snmp.get("43.11.1.1.9.1.2")
	        [int]$yellowpercentremaining = ($yellowcurrentvolume / $yellowtonervolume) * 100
	        $cyantonervolume = $snmp.get("43.11.1.1.8.1.3")
	        $cyancurrentvolume = $snmp.get("43.11.1.1.9.1.3")
	        [int]$cyanpercentremaining = ($cyancurrentvolume / $cyantonervolume) * 100
	        $magentatonervolume = $snmp.get("43.11.1.1.8.1.4")
	        $magentacurrentvolume = $snmp.get("43.11.1.1.9.1.4")
	        [int]$magentapercentremaining = ($magentacurrentvolume / $magentatonervolume) * 100
	        $statustree = $snmp.gettree("43.18.1.1.8")
	        $status = $statustree|? {$_ -notlike "print*"} #status, including low ink warnings
	        $status = $status|? {$_ -notlike "*bypass*"}

	        $name = $snmp.get(".1.3.6.1.2.1.1.5.0")
	        if ($name -notlike "PX*"){$name = $p.name}	
        } 
        elseif (($printertype -like "*Officejet 6700*")) {
            $managed = $true
	        
            $toner_volume = $snmp.GetTree("43.11.1.1.8")
            $toner_current_volume = $snmp.GetTree("43.11.1.1.9")
            $magentatonervolume = $toner_volume[1,3]
	        $magentacurrentvolume = $toner_current_volume[1,3]
	        [int]$magentapercentremaining = ($magentacurrentvolume / $magentatonervolume) * 100
	        $blacktonervolume = $toner_volume[1,0]
	        $blackcurrentvolume = $toner_current_volume[1,0]
	        [int]$blackpercentremaining = ($blackcurrentvolume / $blacktonervolume) * 100
	        $yellowtonervolume = $toner_volume[1,1]
	        $yellowcurrentvolume = $toner_current_volume[1,1]
	        [int]$yellowpercentremaining = ($yellowcurrentvolume / $yellowtonervolume) * 100
	        $cyantonervolume = $toner_volume[1,2]
	        $cyancurrentvolume = $toner_current_volume[1,2]
	        [int]$cyanpercentremaining = ($cyancurrentvolume / $cyantonervolume) * 100

	        $statustree = $snmp.gettree("43.18.1.1.8")
	        $status = $statustree|? {$_ -notlike "print*"} #status, including low ink warnings
	        $status = $status|? {$_ -notlike "*bypass*"}

	        $name = $snmp.get(".1.3.6.1.2.1.1.5.0")
	        if ($name -notlike "PX*"){$name = $p.name}	
        }

        if ($managed) {
            write "<td>"|add-content $outfile
            write ("<b>" + $p.description + "</b>")|add-content $outfile
            write "</td>"|add-content $outfile
            write "<td>"|add-content $outfile
            write ("<a style='text-decoration:none;font-weight:bold;' href=http://" + $p.value + " target='_new'> " + $name + "</a>")|add-content $outfile
            write "</td>"|add-content $outfile
            write "<td>"|add-content $outfile
            write ("<br>" + $printertype + "<br>")|add-content $outfile
            write "</td>"|add-content $outfile
            
            write "<td>"|add-content $outfile
            if ($blackpercentremaining -gt 49){write "<b style='font-size:110%;color:green;'>",$blackpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($blackpercentremaining -gt 24) -and ($blackpercentremaining -le 49)){write "<b style='font-size:110%;color:#40BB30;'>",$blackpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($blackpercentremaining -gt 10) -and ($blackpercentremaining -le 24)){write "<b style='font-size:110%;color:orange;'>",$blackpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($blackpercentremaining -ge 0) -and ($blackpercentremaining -le 10)){write "<b style='font-size:110%;color:red;'>",$blackpercentremaining,"</b>%<br>"|add-content $outfile}
            write "</td>"|add-content $outfile
            
            write "<td>"|add-content $outfile
            if ($cyanpercentremaining -gt 49){write "<b style='font-size:110%;color:green;'>",$cyanpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($cyanpercentremaining -gt 24) -and ($cyanpercentremaining -le 49)){write "<b style='font-size:110%;color:#40BB30;'>",$cyanpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($cyanpercentremaining -gt 10) -and ($cyanpercentremaining -le 24)){write "<b style='font-size:110%;color:orange;'>",$cyanpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($cyanpercentremaining -ge 0) -and ($cyanpercentremaining -le 10)){write "<b style='font-size:110%;color:red;'>",$cyanpercentremaining,"</b>%<br>"|add-content $outfile}
            write "</td>"|add-content $outfile
            
            write "<td>"|add-content $outfile
            if ($magentapercentremaining -gt 49){write "<b style='font-size:110%;color:green;'>",$magentapercentremaining,"</b>%<br>"|add-content $outfile}
            if (($magentapercentremaining -gt 24) -and ($magentapercentremaining -le 49)){write "<b style='font-size:110%;color:#40BB30;'>",$magentapercentremaining,"</b>%<br>"|add-content $outfile}
            if (($magentapercentremaining -gt 10) -and ($magentapercentremaining -le 24)){write "<b style='font-size:110%;color:orange;'>",$magentapercentremaining,"</b>%<br>"|add-content $outfile}
            if (($magentapercentremaining -ge 0) -and ($magentapercentremaining -le 10)){write "<b style='font-size:110%;color:red;'>",$magentapercentremaining,"</b>%<br>"|add-content $outfile}
            write "</td>"|add-content $outfile
            
            write "<td>"|add-content $outfile
            if ($yellowpercentremaining -gt 49){write "<b style='font-size:110%;color:green;'>",$yellowpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($yellowpercentremaining -gt 24) -and ($yellowpercentremaining -le 49)){write "<b style='font-size:110%;color:#40BB30;'>",$yellowpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($yellowpercentremaining -gt 10) -and ($yellowpercentremaining -le 24)){write "<b style='font-size:110%;color:orange;'>",$yellowpercentremaining,"</b>%<br>"|add-content $outfile}
            if (($yellowpercentremaining -ge 0) -and ($yellowpercentremaining -le 10)){write "<b style='font-size:110%;color:red;'>",$yellowpercentremaining,"</b>%<br>"|add-content $outfile}
            write "</td>"|add-content $outfile
            
            write "<td>"|add-content $outfile
            if ($status.length -gt 0)
                {
                write ("<b>"+$status+"</b>" + "<br><br>")|add-content $outfile
                #Send error to monitoring system
                if ($status -like '*toner is out*')
                    {
                       $args = @()
                       $args += ("-CI_Name",$name)   
                       $args += ("-Err_Description",'"' + $status + '"')   
                       Invoke-Expression "$SoapMonitor $args" 
                    }
                elseif ($status -like '*waste toner container is full.*')
                    {
                       $args = @()
                       $args += ("-CI_Name",$name)   
                       $args += ("-Err_Description",'"' + "Waste toner container is full" + '"')   
                       $args += ("-Ins",'"' + "Please check waste toner container" + '"')   
                       Invoke-Expression "$SoapMonitor $args" 
                    }
                }
            else
                {write "Operational<br><br>"|add-content $outfile}
            write "</td>"|add-content $outfile
        }
        write "</tr>"|add-content $outfile
 
    }
}

write "</TABLE>"|add-content $outfile 

write "<h3>",$a,"</h3>"|add-content $outfile 

# Work is now complete, publish file to the final name
Move-Item -Path $outfile -Destination $outfile_final -Force

# uncomment the line to automatically open the html report in your default browser (when running locally) 
#&$outfile_final