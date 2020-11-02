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
        $this.Links = @{
            "How to configure applications for Azure Rights Management" = "https://docs.microsoft.com/en-us/azure/information-protection/configure-applications"
            "Compliance Center - Information Protection"               = "https://compliance.microsoft.com/informationprotection"
            "Compliance Manager - IP Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Information%20protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
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