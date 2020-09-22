using module "..\MCCA.psm1"

class Audit101 : MCCACheck {
    <#
    

    #>

    Audit101() {
        $this.Control = "Audit-101"
        $this.ParentArea = "Discovery & Response"
        $this.Area = "Audit"
        $this.Name = "Enable Auditing in Office 365"
        $this.PassText = "Your organisation has enabled auditing for your Office 365 tenant"
        $this.FailRecommendation = "Your organization should enable auditing for your Office 365 tenant"
        $this.Importance = "Your organization should enable auditing for your Office 365 tenant. When audit log search in the Security & Compliance Center is turned on, user and admin activity from your organization is recorded in the audit log and retained for 90 days, and up to one year depending on the license assigned to users."
        $this.ExpandResults = $True
        $this.ItemName = "Configuration"
        $this.DataType = "Setting"
        $this.Links = @{
            "How to search Audit Log"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance"
            "Advanced Audit"                       = "https://docs.microsoft.com/en-us/microsoft-365/compliance/advanced-audit"
            "Compliance Center - Audit Log search" = "https://compliance.microsoft.com/auditlogsearch"
            "Compliance Manager - Audit Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Audit%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22OutOfScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        if ($Config["GetAdminAuditLogConfig"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            $Auditconfiguration = $Config["GetAdminAuditLogConfig"]
            $ConfigObject = [MCCACheckConfig]::new()
            $ConfigObject.Object = "Configuration"
            $ConfigObject.ConfigItem = "Auditing in Office 365"
            
            # Determine if UnifiedAuditLogIngestionEnabled is true in Audit Configuration
            If ($($Auditconfiguration.UnifiedAuditLogIngestionEnabled) -eq $true) {
                $ConfigObject.ConfigData = "Enabled"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
            } 
            Else {
                $ConfigObject.ConfigData = "Disabled"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                $ConfigObject.RemediationAction = "Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled " + "$" + "true"
                Write-Host "$(Get-Date) Generating Remediation Action to enable Auditing" -ForegroundColor Yellow

            }

            $this.AddConfig($ConfigObject)
            $ConfigObjectList += $ConfigObject
            $hasRemediation = $this.Config | Where-Object { $_.RemediationAction -ne ''}
            if ($($hasremediation.count) -gt 0)
            {
                $this.MCCARemediationInfo = New-Object -TypeName MCCARemediationInfo -Property @{
                    RemediationAvailable = $True
                    RemediationText      = "You need to connect to Exchange Online Center PowerShell to execute the below commands. Please follow steps defined in <a href = 'https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps'> Connect to Exchange Online Center PowerShell</a>."
                }
            }
            $this.Completed = $True
        }
        
    }

}