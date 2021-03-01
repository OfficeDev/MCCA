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
                "Overview of sensitivity labels "                           = "https://aka.ms/mcca-ip-docs-sensitivity-labels"
                "How to apply a sensitivity label to content automatically" = "https://aka.ms/mcca-ip-docs-action-apply-sensitivity-labels"
                "Compliance Center - Information Protection"               = "https://aka.ms/mcca-gcch-ip-compliance-center"
                "Compliance Manager - IP Actions" = "https://aka.ms/mcca-gcch-ip-compliance-manager"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Overview of sensitivity labels "                           = "https://aka.ms/mcca-ip-docs-sensitivity-labels"
                "How to apply a sensitivity label to content automatically" = "https://aka.ms/mcca-ip-docs-action-apply-sensitivity-labels"
                "Compliance Center - Information Protection"               = "https://aka.ms/mcca-dod-ip-compliance-center"
                "Compliance Manager - IP Actions" = "https://aka.ms/mcca-dod-ip-compliance-manager"
            } 
        }else
        {
        $this.Links = @{
            "Overview of sensitivity labels "                           = "https://aka.ms/mcca-ip-docs-sensitivity-labels"
            "How to apply a sensitivity label to content automatically" = "https://aka.ms/mcca-ip-docs-action-apply-sensitivity-labels"
            "Compliance Center - Information Protection"               = "https://aka.ms/mcca-ip-compliance-center"
            "Compliance Manager - IP Actions" = "https://aka.ms/mcca-ip-compliance-manager"
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