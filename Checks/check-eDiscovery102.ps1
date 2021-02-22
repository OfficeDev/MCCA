using module "..\MCCA.psm1"

class eDiscovery102 : MCCACheck {
    <#
    

    #>

    eDiscovery102() {
        $this.Control = "eDiscovery-102"
        $this.ParentArea = "Discovery & Response"
        $this.Area = "eDiscovery"
        $this.Name = "Use Advanced eDiscovery Cases to Support Legal Investigations"
        $this.PassText = "Your organization is using Advanced eDiscovery cases to support legal investigations"
        $this.FailRecommendation = "Your organization needs to review (or set up) Advanced eDiscovery cases"
        $this.Importance = "Your organization should use Advanced eDiscovery to manage the end-to-end workflow to preserve, collect, review, analyze, and export content that's responsive to your organization's internal and external investigations."
        $this.ExpandResults = $True
        $this.ItemName = "Case Name"
        $this.DataType = "Case Status"
        if($this.ExchangeEnvironmentNameForCheck -eq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Get started with Advanced eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-advanced-ediscovery?view=o365-worldwide"
                "Compliance Center - Advanced eDiscovery"                       = "https://compliance.microsoft.us/advancedediscovery"
                "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
                "Compliance Manager - eDiscovery" = "https://compliance.apps.mil/compliancemanager?filter=%7B%22Solution%22:%5B%22eDiscovery%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -eq "O365USGovDoD") 
        {
            $this.Links = @{
                "Get started with Advanced eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-advanced-ediscovery?view=o365-worldwide"
                "Compliance Center - Advanced eDiscovery"                       = "https://compliance.apps.mil/advancedediscovery"
                "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
                "Compliance Manager - eDiscovery" = "https://compliance.apps.mil/compliancemanager?filter=%7B%22Solution%22:%5B%22eDiscovery%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            }
        }else
        {
        $this.Links = @{
            "Get started with Advanced eDiscovery"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-advanced-ediscovery?view=o365-worldwide"
            "Compliance Center - Advanced eDiscovery"                       = "https://compliance.microsoft.com/advancedediscovery"
            "eDiscovery in Microsoft 365" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/ediscovery?view=o365-worldwide"
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        if ($Config["GetComplianceCase"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            
            $CasesPresent= $false
            $activecasepresent = $false
            ForEach ($CasesDefined in $Config["GetComplianceCase"]|Where-Object{$_.CaseType -eq "AdvancedEdiscovery"}) { 
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
    }}
