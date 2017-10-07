# PowEnum

Penetration testers commonly enumerate AD data – providing domain situational awareness and helping to identify soft targets.  PowEnum helps automate the cartological view of your target domain.

PowEnum executes common PowerSploit Powerview functions and combines the output into a spreadsheet for easy analysis. All network traffic is only sent to the DC(s). PowEnum also leverages PowerSploit Get-GPPPassword and Harmj0y's ASREPRoast.

#### Syntax Examples:
  - Invoke-PowEnum
  - Invoke-PowEnum -FQDN test.domain.com
  - Invoke-PowEnum -Mode SYSVOL
  - Invoke-PowEnum -Credential test.domain.com\username -FQDN test.domain.com -Mode Special

### Running PowEnum From Non-Domain Joined System
There are two choices. The first uses the runas command (this must be executed prior to using PowEnum). The second leverages the Invoke-UserImpersonation function in Powerview.
1) runas /netonly /user:test.domain.com\username powershell.exe
2) Invoke-PowEnum -Credential test.domain.com\username -FQDN test.domain.com

### Modes

| Mode | Enumerates | 
| ------ | ------ |
| Basic | Domain Admins<br>Enterprise Admins<br>Built-In Admins<br>DC Local Admins <br>Domain Users<br>Domain Groups<br>Schema Admin<br>Account Operators<br>Backup Operators<br>Print Operators<br>Server Operators<br>Group Policy Creators Owners<br>Cryptographic Operators<br>AD Group Managers<br>AdminCount=1<br><br> All [DC Aware] Net Sessions<br>Domain Controllers<br>Domain Computer IPs<br>Domain Computers<br>Subnets<br>DNSRecords<br>WinRM Enabled Hosts<br>Potential Fileservers |
| Roasting | Kerberoast Service Accounts (Accounts w/ SPN)<br>ASREPRoast User Accounts (No Preauth Req) |
| Special | Disabled Accounts<br>Password Not Required<br>Password Doesn't Expire<br>Password Doesn't Expire & Not Required <br>Smartcard Required |
| SYSVOL | Group Policy Passwords<br>SYSVOL Script Files (potential hardcoded credentials) |
| Forest | Domain Trusts<br>Foreign [Domain] Users<br>Foreign [Domain] Group Members |
| LargeEnv | Basic Enumeration without:<br>Get-DomainUser<br>Get-DomainGroup<br>Get-DomainComputer|

*DC Local Admins might be different from built-in Administrators when an RODC is in use or there are replication issues.

### Detection
  - This enumeration will generate suspicious traffic between the PowEnum system and the target DC(s). If there are security products watching traffic to the DC(s) (i.e. Microsoft ATA) PowEnum will likely get flagged. For more reading about what ATA is detecting and not detecting:<br>https://media.defcon.org/DEF%20CON%2025/DEF%20CON%2025%20presentations/DEFCON-25-Chris-Thompson-MS-Just-Gave-The-Blue-Teams-Tactical-Nukes-UPDATED.pdf
  - Kerberoasting detection techniques are highlighted in these articles:<br>https://adsecurity.org/?p=3458<br>https://adsecurity.org/?p=3513 

### Mitigations
| Mode | Mitigations |
| ------ | ------ |
| Basic | Net Cease - Hardening Net Session Enumeration<br>https://gallery.technet.microsoft.com/Net-Cease-Blocking-Net-1e8dcb5<br>SAMRi10 - Hardening SAM Remote Access in Windows 10/Server 2016<br>https://gallery.technet.microsoft.com/SAMRi10-Hardening-Remote-48d94b5b<br>Active Directory: Controlling Object Visibility<br>https://social.technet.microsoft.com/wiki/contents/articles/29558.active-directory-controlling-object-visibility-list-object-mode.aspx<br>http://windowsitpro.com/active-directory/hiding-active-directory-objects-and-attributes |
| Roasting | Kerberoasst mitigations revolve around using strong passwords or GMSA for affected accounts<br>https://adsecurity.org/?p=2293<br>ASREPRoast mitigations revolve around using strong passwords or not checking "‘Do Not Require Kerberos Preauthentication" |
| Special | See Basic  |
| SYSVOL | GPP Password Files - Install KB2962486  and remove affected xml files (https://adsecurity.org/?p=2288)<br> SYSVOL Scripts - Monitor for changes to SYSVOL and remove affected files |
| Forest | See Basic |
| LargeEnv | See Basic |
