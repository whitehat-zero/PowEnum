# PowEnum

Executes common PowerSploit Powerview functions and then combines the output into a spreadsheet for easy analysis. All network traffic is only sent to the DC(s).

#### Syntax Examples:
  - Invoke-PowEnum
  - Invoke-PowEnum -URL http://10.0.0.10/PowerView.ps1
  - Invoke-PowEnum -Domain test.com
  - Invoke-PowEnum -Mode Special
  - Invoke-PowEnum -Credential test.domain.com\username -Mode Special

### Modes

| Mode | Enumerates |
| ------ | ------ |
| Basic | Domain Admins<br>Enterprise Admins<br>Built-In Admins<br>DC Local Admins<br>Domain Users<br>Domain Groups<br>All [DC Aware] Net Sessions<br>Domain Controllers<br>Domain Computer IPs<br>Domain Computers<br>Subnets<br>DNSRecords<br>WinRM Enabled Hosts |
| Roasting | Kerberoast Service Accounts<br>ASREPRoast User Accounts |
| LargeEnv | Basic Enumeration without Get-DomainUser/Get-DomainGroup/Get-DomainComputer |
| Special | Disabled Accounts<br>Password Not Required<br>Password Doesn't Expire<br>Password Doesn't Expire & Not Required <br>Smartcard Required |

### Detection
This enumeration is noisy if there are security products watching traffic to the DC(s) (i.e. Microsoft ATA)

### Mitigations
Net Cease - Hardening Net Session Enumeration
https://gallery.technet.microsoft.com/Net-Cease-Blocking-Net-1e8dcb5b
SAMRi10 - Hardening SAM Remote Access in Windows 10/Server 2016
https://gallery.technet.microsoft.com/SAMRi10-Hardening-Remote-48d94b5b

## TODO
- Check Credential For FQDN
- Create ASREPRoast Url Parameter
- Add Following Groups To Group Enumeration
  - Schema Admin
  - Account Operators
  - Backup Operators
  - Print Operators
  - Server Operators
  - Domain Controllers
  - Read-only Domain Controllers
  - Group Policy Creators Owners
  - Cryptographic Operators
- Add GPP and SYSVOL searching (.vbs, .bat, .cmd)

