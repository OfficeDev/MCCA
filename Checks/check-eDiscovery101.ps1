using module "..\MCCA.psm1"

class eDiscovery101 : MCCACheck {
    <#
    

    #>

    eDiscovery101() {
        $this.Control = "eDiscovery-101"
        $this.ParentArea = "Discovery & Response"
        $this.Area = "eDiscovery"
        $this.Name = "Use Core eDiscovery Cases to Support Legal Investigations"
        $this.PassText = "Your organization is using Core eDiscovery cases to support legal investigations"
        $this.FailRecommendation = "Your organization needs to review (or set up) Core eDiscovery cases"
        $this.Importance = "Your organization should use Core eDiscovery cases to identify, hold, and export content found in Exchange Online mailboxes, Microsoft 365 Groups, Microsoft Teams, SharePoint Online and OneDrive for Business sites, and Skype for Business conversations, and Yammer teams."
        $this.ExpandResults = $True
        $this.ItemName = "Case Name"
        $this.DataType = "Case Status"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Get started with Core eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-core-ediscovery?view=o365-worldwide"
                "Compliance Center - Core eDiscovery"                       = "https://compliance.microsoft.us/classicediscovery"
                "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
                "Compliance Manager - eDiscovery" = "https://compliance.microsoft.us/compliancemanager?filter=%7B%22Solution%22:%5B%22eDiscovery%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Get started with Core eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-core-ediscovery?view=o365-worldwide"
                "Compliance Center - Core eDiscovery"                       = "https://compliance.apps.mil/classicediscovery"
                "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
                "Compliance Manager - eDiscovery" = "https://compliance.apps.mil/compliancemanager?filter=%7B%22Solution%22:%5B%22eDiscovery%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            }
        }else
        {
        $this.Links = @{
            "Get started with Core eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-core-ediscovery?view=o365-worldwide"
            "Compliance Center - Core eDiscovery"                       = "https://compliance.microsoft.com/classicediscovery"
            "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        if ($Config["GetComplianceCaseCore"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            
            $CasesPresent= $false
            $activecasepresent = $false
            ForEach ($CasesDefined in $Config["GetComplianceCaseCore"]) { 
                $Cases = $CasesDefined 
                $CasesPresent= $true
                
                if($($Cases.Status) -eq "Active")
                {
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.ConfigItem = "$($Cases.Name)"
                    $ConfigObject.ConfigData = "$($Cases.Status)"
                    $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                    $this.AddConfig($ConfigObject)
                    $ConfigObjectList += $ConfigObject
                    $activecasepresent= $true
                }
                $CasesPresent= $true
            }
            if(($activecasepresent -eq $false)  -and ($CasesPresent -eq $true))
            {
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.ConfigItem = "No active case"
                    $ConfigObject.ConfigData = ""
                    $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
                    $this.AddConfig($ConfigObject)
                    $ConfigObjectList += $ConfigObject
            }
            elseif($CasesPresent -eq $false)
            {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.ConfigItem = "No eDiscovery cases found"
                $ConfigObject.ConfigData = ""
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                $this.AddConfig($ConfigObject)
                $ConfigObjectList += $ConfigObject
            }
            
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
