# PowEnum

Executes common PowerSploit Powerview functions and then combines the output into a spreadsheet.

#### Syntax Examples:
  - Invoke-PowEnum
  - Invoke-PowEnum -URL http://10.0.0.10/PowerView.ps1
  - Invoke-PowEnum -Domain test.com
  - Invoke-PowEnum -Mode Special
  - Invoke-PowEnum -Credential (Get-Credential) -Mode Special

### Modes

| Mode | Enumerates |
| ------ | ------ |
| Basic | Domain Admins<br>Enterprise Admins<br>Built-In Admins<br>DC Local Admins<br>Domain Users<br>Domain Groups<br>All [DC Aware] Net Sessions<br>Domain Controllers<br>Domain Computer IPs<br>Domain Computers<br>Subnets<br>DNSRecords<br>WinRM Enabled Hosts |
| Roasting | Kerberoast Service Accounts<br>ASREPRoast User Accounts |
| LargeEnv | Disabled Account<br>Password Not Required<br>Password Doesn't Expire<br>Password Doesn't Expire & Not Required <br>Smartcard Required |
