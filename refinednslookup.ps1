
#redefined nslook
#written by Peter McLane @GIS JCI

#Used to automate nslookup command 
$xlsx = New-Object -ComObject excel.application
$wb = $xlsx.workbooks.Open("C:\Users\*****\OneDrive - Johnson Controls\Documents\WAF_Project\external_ips.xlsx")
$sheet = $wb.Sheets.Item(1)
$usedrange = $sheet.UsedRange




#fix output so it is easier to read??
#number of rows in spreadsheet
$start = 1
$count = 1189
while($start -lt $count){
    #Write-Host Row: $start
    #echo $usedrange.cells($start,1).value2
	(nslookup $usedrange.cells($start,1).value2 | Select-Object -ExpandProperty "") -join
    #Start-Sleep -s 1
    #echo------------------------------------------------------------------------------------------------------------------------------- 
	$start++
}


$wb.Close()


#export-excel to print results to spreadsheet
