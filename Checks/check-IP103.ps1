using module "..\MCCA.psm1"

class IP103 : MCCACheck {
    <#
    

    #>

    IP103() {
        $this.Control = "IP-103"
        $this.ParentArea = "Microsoft Information Protection"
        $this.Area = "Information Protection"
        $this.Name = "Use IRM for Exchange Online"
        $this.PassText = "Your organization has enabled IRM for Exchange Online"
        $this.FailRecommendation = "Your organization should enable IRM for Exchange Online"
        $this.Importance = "Your organization should enable and use Azure Information Protection for Exchange Online. This configuration lets Exchange provide protection solutions, such as mail flow rules, data loss prevention policies that contain sets of conditions to filter email messages and take actions, and protection rules for Outlook clients."
        $this.ExpandResults = $True
        $this.ItemName = "IRM Configuration"
        $this.DataType = "Setting"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "How to configure applications for Azure Rights Management" = "https://aka.ms/mcca-ip-docs-action-ARM"
                "Compliance Center - Information Protection"               = "https://aka.ms/mcca-gcch-ip-compliance-center"
                "Compliance Manager - IP Actions" = "https://aka.ms/mcca-gcch-ip-compliance-manager"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "How to configure applications for Azure Rights Management" = "https://aka.ms/mcca-ip-docs-action-ARM"
                "Compliance Center - Information Protection"               = "https://aka.ms/mcca-dod-ip-compliance-center"
                "Compliance Manager - IP Actions" = "https://aka.ms/mcca-dod-ip-compliance-manager"
            } 
        }else
        {
        $this.Links = @{
            "How to configure applications for Azure Rights Management" = "https://aka.ms/mcca-ip-docs-action-ARM"
            "Compliance Center - Information Protection"               = "https://aka.ms/mcca-ip-compliance-center"
            "Compliance Manager - IP Actions" = "https://aka.ms/mcca-ip-compliance-manager"
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        if ($Config["GetIRMConfiguration"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            $IRMconfiguration = $Config["GetIRMConfiguration"]
            $ConfigObject = [MCCACheckConfig]::new()
            $ConfigObject.Object = "IRM Configuration"
            $ConfigObject.ConfigItem = "AzureRMSLicensingEnabled"
            $ConfigObject.ConfigData = $IRMconfiguration.AzureRMSLicensingEnabled

            # Determine if AzureRMSLicensingEnabled is true in IRM Configuration
            If ($IRMconfiguration.AzureRMSLicensingEnabled -eq $true) {
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
            } 
            Else {
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
            }
            $ConfigObjectList += $ConfigObject
            $this.AddConfig($ConfigObject)
        

            $hasRemediation = $this.Config | Where-Object { $_.RemediationAction -ne ''}
            if ($($hasremediation.count) -gt 0)
            {
                $this.MCCARemediationInfo = New-Object -TypeName MCCARemediationInfo -Property @{
                    RemediationAvailable = $True
                    RemediationText      = "You need to connect to Security & Compliance Center PowerShell to execute the below commands. Please follow steps defined in <a href = 'https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'> Connect to Security & Compliance Center PowerShell</a>."
                }
            }
            $this.Completed = $True
        }
      
    }

}