using module "..\MCCA.psm1"

class IP102 : MCCACheck {
    <#
    
         
    #>

    IP102() {
        $this.Control = "IP-102"
        $this.ParentArea = "Microsoft Information Protection"
        $this.Area = "Information Protection"
        $this.Name = "Auto-apply client side sensitivity labels"
        $this.PassText = "Your organization is using auto-apply client side sensitivity labels"
        $this.FailRecommendation = "Your organization should use client side sensitivity labels"
        $this.Importance = "Your organization should automatically apply client side sensitivity labels based on sensitive information types or other criteria. Microsoft recommends that automatic labeling be implemented to decrease reliance on users for correct classification."
        $this.ExpandResults = $True
        $this.ItemName = "Labels"
        $this.DataType = "Remarks"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Overview of sensitivity labels "                           = "https://docs.microsoft.com/en-us/microsoft-365/compliance/sensitivity-labels?view=o365-worldwide"
                "How to apply a sensitivity label to content automatically" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/apply-sensitivity-label-automatically?view=o365-worldwide"
                "Compliance Center - Information Protection"               = "https://compliance.microsoft.us/informationprotection"
                "Compliance Manager - IP Actions" = "https://compliance.microsoft.us/compliancemanager?filter=%7B%22Solution%22:%5B%22Microsoft%20Information%20Protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Overview of sensitivity labels "                           = "https://docs.microsoft.com/en-us/microsoft-365/compliance/sensitivity-labels?view=o365-worldwide"
                "How to apply a sensitivity label to content automatically" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/apply-sensitivity-label-automatically?view=o365-worldwide"
                "Compliance Center - Information Protection"               = "https://compliance.apps.mil/informationprotection"
                "Compliance Manager - IP Actions" = "https://compliance.apps.mil/compliancemanager?filter=%7B%22Solution%22:%5B%22Microsoft%20Information%20Protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22%5D%7D&viewid=ImprovementActions"
            } 
        }else
        {
        $this.Links = @{
            "Overview of sensitivity labels "                           = "https://docs.microsoft.com/en-us/microsoft-365/compliance/sensitivity-labels?view=o365-worldwide"
            "How to apply a sensitivity label to content automatically" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/apply-sensitivity-label-automatically?view=o365-worldwide"
            "Compliance Center - Information Protection"               = "https://compliance.microsoft.com/informationprotection"
            "Compliance Manager - IP Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Information%20protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {
        if ($Config["GetLabel"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            $AutoApplyExist = $false
        
            ForEach ($LabelsDefined in $Config["GetLabel"]) { 
                $Labels = $LabelsDefined 
    
                if ($($Labels.Conditions)) {
                    if ($($Labels.Disabled) -eq $false) {
                        $Workload = $Labels.Workload

                        if ((((@($Workload) -like 'Exchange').Count -lt 1)) -and (((@($Workload) -like 'SharePoint').Count -lt 1))) {
                            $ConfigObject = [MCCACheckConfig]::new()
                            $ConfigObject.ConfigItem = "$($Labels.DisplayName)"
                            $ConfigObject.ConfigData = "Enabled Workloads: Office Apps"
                            $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
                            $ConfigObjectList += $ConfigObject
                            $AutoApplyExist = $true
                            $this.AddConfig($ConfigObject)


                        }
                        else {
                            $ConfigObject = [MCCACheckConfig]::new()
                            $ConfigObject.ConfigItem = "$($Labels.DisplayName)"
                            $ConfigObject.ConfigData = "Only Enabled Workload: $($Labels.Workload)"
                            $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                            $ConfigObjectList += $ConfigObject
                            $this.AddConfig($ConfigObject)

                        }
                    }
                    else {
                        $ConfigObject = [MCCACheckConfig]::new()
                        $ConfigObject.ConfigItem = "$($Labels.DisplayName)"
                        $ConfigObject.ConfigData = "Label is not enabled"
                        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                        $ConfigObjectList += $ConfigObject
                        $this.AddConfig($ConfigObject)

                    }
           
                }

            }


            If ($AutoApplyExist -eq $False) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "No Active Policy defined"
                $ConfigObject.ConfigItem = "No Auto Apply Policy"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")      
                $ConfigObjectList += $ConfigObject      
                $this.AddConfig($ConfigObject)
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