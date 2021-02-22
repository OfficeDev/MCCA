Remove-Module MCCAPreview -ErrorAction SilentlyContinue
Remove-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
Unblock-File ".\*"
Unblock-File ".\Checks\*"
Unblock-File ".\Outputs\*"
Unblock-File ".\Remediation\*"
Unblock-File ".\Utilities\*"

Import-Module .\MCCA.psm1


#Get-MCCAReport -Geo @("nam") -Solution @("num")

#Get-MCCAReport -ExchangeEnvironmentName O365USGovGCCHigh




