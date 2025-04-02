#Cole patterson
$Version = "1.3"
$Date = "3/26/2025"

#Starting Spooler
net Stop Spooler
net start Spooler

#Gets Print Server Printers
$Print = Get-Printer -ComputerName "Print" | where Shared -eq $true 

#Gets Local Printers  #Was not working on Script server
$AvailablePrinters = Get-WmiObject -ClassName Win32_Printer

foreach ($item in $Print.name) {
    #Sets Default printer
	Try{($AvailablePrinters | Where-Object -FilterScript {$_.Name -eq "$item"}).SetDefaultPrinter() | Out-Null}
	Catch{write-host "Offline"}
	
	Write-host "$item"
    Start-Sleep 5
	Start-Process -FilePath "C:\Script\TestPage.docx" -Verb Print
	#Start-Process -FilePath "C:\Script\TestPage.docx" -Verb PrintTo "\\Print\$item" -PassThru
	Write-host "Printed on $item"
}
