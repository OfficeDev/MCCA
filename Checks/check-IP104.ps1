using module "..\MCCA.psm1"

class IP104 : MCCACheck {
    <#
    
         
    #>

    IP104() {
        $this.Control = "IP-104"
        $this.ParentArea = "Microsoft Information Protection"
        $this.Area = "Information Protection"
        $this.Name = "Create service side labelling policies"
        $this.PassText = "Your organization is using service side labeling policies"
        $this.FailRecommendation = "Your organization should use service side labeling policies"
        $this.Importance = "Your organization should setup and create service side labelling policies . This will help categorize your most important data so that you can effectively protect it from illicit access, and will help make it easier to investigate discovered breaches."
        $this.ExpandResults = $True
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ObjectType = "Auto Labelling Policy"
        $this.ItemName = "Label"
        $this.DataType = "Remarks"
        $this.Links = @{
            "Learn more about configuring classifications for SharePoint Online" = "https://go.microsoft.com/fwlink/?linkid=2071059"
            "Compliance Center - Information Protection"                        = "https://compliance.microsoft.com/informationprotection"
            "Compliance Manager - IP Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Information%20protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {

        if ($Config["GetAutoSensitivityLabelPolicy"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $AutoApplyExist = $false
            $isExchangeCovered = $false
            $isSharePointCovered = $false
            $isOneDriveCovered = $false
            ForEach ($AutoPolicyDefined in $Config["GetAutoSensitivityLabelPolicy"]) { 
                $AutoPolicy = $AutoPolicyDefined 
                $AutoApplyExist = $true         
                #Validate if Auto labelling policies are enabled across all workloads 
                 
                if ($($AutoPolicy.Mode) -eq "Disable") {
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.Object = "$($AutoPolicy.Name)"
                    $ConfigObject.ConfigItem = "$($AutoPolicy.ApplySensitivityLabel)"
                    $ConfigObject.ConfigData = "<B>Policy is not enabled  </B> "
                    $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")
                    $this.AddConfig($ConfigObject)
                }
                else {
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.Object = "$($AutoPolicy.Name)"
                    $ConfigObject.ConfigItem = "$($AutoPolicy.ApplySensitivityLabel)"
                    $ConfigData = $null
                    if ( ($null -ne $($AutoPolicy.ExchangeLocation)  ) -and ($null -ne $($AutoPolicy.SharePointLocation) ) -and ($null -ne $($AutoPolicy.OneDriveLocation) )) {
                        $ConfigData = "<B>Exchange User/Groups:</B> $($AutoPolicy.ExchangeLocation) <BR>"
                        $ConfigData += "<B>SharePoint Sites:</B> $($AutoPolicy.SharePointLocation) <BR>"
                        $ConfigData += "<B>OneDrive Accounts:</B> $($AutoPolicy.OneDriveLocation) <BR>"
                        $ConfigObject.ConfigData = $ConfigData
                        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
                        $this.AddConfig($ConfigObject)
                        $isExchangeCovered = $true
                        $isSharePointCovered = $true
                        $isOneDriveCovered = $true
                   
                    }
                    else {
                        $ConfigObject = [MCCACheckConfig]::new()
                        $ConfigObject.Object = "$($AutoPolicy.Name)"
                        $ConfigObject.ConfigItem = "$($AutoPolicy.ApplySensitivityLabel)"
                        $ConfigData = $null
                        if ( ($null -ne $($AutoPolicy.ExchangeLocation)  ) -and ($null -ne $($AutoPolicy.SharePointLocation) ) -and ($null -ne $($AutoPolicy.OneDriveLocation) )) {
                            $ConfigData = "<B>Exchange User/Groups:</B> $($AutoPolicy.ExchangeLocation) <BR>"
                            $isExchangeCovered = $true
                        }
                        else {
                            $ConfigData = "<B>Exchange User/Groups:</B> Not Enabled <BR>"
                        }
                        if ($null -ne $($AutoPolicy.SharePointLocation)  ) {
                            $ConfigData += "<B>SharePoint Sites:</B> $($AutoPolicy.SharePointLocation) <BR>"
                            $isSharePointCovered = $true
                        }
                        else {
                            $ConfigData += "<B>SharePoint Sites:</B> Not Enabled <BR>"
                        }
                        if ($null -ne $($AutoPolicy.OneDriveLocation)  ) {
                            $ConfigData += "<B>OneDrive Accounts:</B> $($AutoPolicy.OneDriveLocation) <BR>"
                            $isOneDriveCovered = $true
                        }
                        else {
                            $ConfigData += "<B>OneDrive Accounts:</B> Not Enabled <BR>"    

                        }
                        $ConfigObject.ConfigData = $ConfigData
                        $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")
                        $this.AddConfig($ConfigObject)
                    }
                 
                }
            }
            $PartialWorkload = ""
            If ($isExchangeCovered -eq $false) {
                if ($PartialWorkload -eq "") {
                    $PartialWorkload += "Exchange"   
                }
                else {            
                    $PartialWorkload += ",Exchange" 
                } 
            }

            If ($isSharePointCovered -eq $false) {
                if ($PartialWorkload -eq "") {
                    $PartialWorkload += "SharePoint"   
                }
                else {            
                    $PartialWorkload += ",SharePoint" 
                } 
            }

            If ($isOneDriveCovered -eq $false) {
                if ($PartialWorkload -eq "") {
                    $PartialWorkload += "OneDrive"   
                }
                else {            
                    $PartialWorkload += ",OneDrive" 
                } 
            }
            #policy not defined on one or more workload
            If (($PartialWorkload -ne "") -and ($AutoApplyExist -eq $true) ){
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>All workload not covered</B>"
                #$ConfigObject.ConfigItem = $PartialLabel
                $ConfigData = "<B>Affected Workloads:</B>$PartialWorkload <BR>"
                $ConfigObject.ConfigData = $ConfigData
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                $this.AddConfig($ConfigObject)
            }
            If ($AutoApplyExist -eq $False) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<b>No Auto Labeling Policy Defined</b>"
                $ConfigObject.ConfigItem = ""
                $ConfigData = "<B>Affected Workloads:</B>Exchange, SharePoint, OneDrive"
                $ConfigObject.ConfigData = $ConfigData
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
                $this.AddConfig($ConfigObject)
            }
            $hasRemediation = $this.Config | Where-Object { $_.RemediationAction -ne '' }
            if ($($hasremediation.count) -gt 0) {
                $this.MCCARemediationInfo = New-Object -TypeName MCCARemediationInfo -Property @{
                    RemediationAvailable = $True
                    RemediationText      = "You need to connect to Security & Compliance Center PowerShell to execute the below commands. Please follow steps defined in <a href = 'https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'> Connect to Security & Compliance Center PowerShell</a>."
                }
            }
            $this.Completed = $True
        }
    }

}