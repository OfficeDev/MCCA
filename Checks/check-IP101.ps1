using module "..\MCCA.psm1"

class IP101 : MCCACheck {
    <#
    
         
    #>

    IP101() {
        $this.Control = "IP-101"
        $this.ParentArea = "Microsoft Information Protection"
        $this.Area = "Information Protection"
        $this.Name = "Create Sensitivity Labels for Sensitive or Critical Data"
        $this.PassText = "Your organization is using sensitivity labels to classify your information"
        $this.FailRecommendation = "Your organization should be using sensitivity labels to classify your information"
        $this.Importance = "Your organization should use sensitivity labels and policies to classify your information in SharePoint Online, OneDrive for Business, and Exchange Online. This helps categorize your most important data and effectively protect it from illicit access; it can also make it easier to investigate discovered breaches."
        $this.ExpandResults = $True
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ObjectType = "Label Policy"
        $this.ItemName = "Labels"
        $this.DataType = "Remarks"
        $this.Links = @{
            "Overview of sensitivity labels "                                     = "https://docs.microsoft.com/en-us/microsoft-365/compliance/sensitivity-labels?view=o365-worldwide"
            "How to configure classifications for your Microsoft 365 environment" = "https://docs.microsoft.com/en-us/microsoft-365/enterprise/infoprotect-configure-classification?view=o365-worldwide"
            "Compliance Center - Information Protection"                         = "https://compliance.microsoft.com/informationprotection"
            "Compliance Manager - IP Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Information%20protection%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22OutOfscope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {
        if (($Config["GetLabel"] -eq "Error") -or ($Config["GetLabelPolicy"] -eq "Error")) {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
        
            $LabelPolicyExists = $False
               
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
                . $UtilityFile.FullName
            }
            $LogFile = $this.LogFile
            $LabelAssociation = Get-LableCalssification -LogFile $LogFile
            $SubLabels = $LabelAssociation.sublabels
            $ParentLabels = $LabelAssociation.parentlabels
            $ParentSubLabelAssociation = $LabelAssociation.parentsublabelassociation
            $ParentNameForSubLabelAssociation = $LabelAssociation.parentnameforsublabelassociation

            $AllGlobalPolicy = 0
            $ScopedGlobalPolicy = 0
            ForEach ($Policies in $Config["GetLabelPolicy"]) {   
                $LabelPolicy = $Policies
                $ConfigObject = [MCCACheckConfig]::new()
                $PolicyParentLabelCount = 0
                $PolicySubLabelCount = 0
                $ParentLabelsDefinedinPolicy = "None"
                $SubLabelsDefinedinPolicy = "None"
                if ($LabelPolicy.Enabled -eq $true) {
                    foreach ($LabelConfigured in $LabelPolicy.Labels) {  
                        if ($SubLabels.containsKey($LabelConfigured)) {
                            $PolicySubLabelCount++
                            if ($SubLabelsDefinedinPolicy -eq "None") {
                                $SubLabelsDefinedinPolicy = $LabelConfigured 
                            }
                            else {
                                $SubLabelsDefinedinPolicy += ", $LabelConfigured"
                            }
                        }
                        if ($ParentLabels.containsKey($LabelConfigured)) {
                            $PolicyParentLabelCount++
                            if ($ParentLabelsDefinedinPolicy -eq "None") {
                                $ParentLabelsDefinedinPolicy = $LabelConfigured 
                            }
                            else {
                                $ParentLabelsDefinedinPolicy += ", $LabelConfigured"
                            }
                        }

                    }
                    $ExchangeLocation = $($LabelPolicy.ExchangeLocation)
                    if ((@($ExchangeLocation) -like 'All').Count -gt 0) {
                        $AllGlobalPolicy++
                    }
                    else {
                        $ScopedGlobalPolicy++
                    }
                    if ($PolicyParentLabelCount -le 9) {
                        $ConfigObject.Object = $LabelPolicy.Name
                        $ConfigObject.ConfigItem = "<B>Parent Labels</B> : $ParentLabelsDefinedinPolicy <BR> <B>Sub Labels</B> : $SubLabelsDefinedinPolicy"
                        $ConfigObject.ConfigData = "<B>Enabled Workload:</B> $($LabelPolicy.ExchangeLocation) $($LabelPolicy.ModernGroupLocation)"
                        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
                        $ConfigObjectList += $ConfigObject 
                    }
                    else {
                        $ConfigObject.Object = "$($LabelPolicy.Name)"
                        $ConfigObject.ConfigItem = "<B>Parent Labels</B> : $ParentLabelsDefinedinPolicy <BR> <B>Sub Labels</B> : $SubLabelsDefinedinPolicy"
                        $ConfigObject.ConfigData = "No of parent Labels is more than 8."
                        $ConfigObject.InfoText = "You currently have $($PolicyParentLabelCount) Global policies defined. We have found 8 parent labels to be optimal for most organizations"
                        $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Fail")
                        $ConfigObjectList += $ConfigObject 
                                             
                    }
                    $LabelPolicyExists = $true
                }
                else {
                    $ConfigObject.Object = "$($LabelPolicy.Name)"
                    $ConfigObject.ConfigItem = "$($LabelPolicy.Labels)"
                    $ConfigObject.ConfigData = " <B>Policy is not enabled </B>"   
                    $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")      
                    $ConfigObjectList += $ConfigObject 
                }
                $this.AddConfig($ConfigObject)
                $LabelPolicyExists = $true


            }

            if ($AllGlobalPolicy -gt 5) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>High Number of Global policies defined<B>"
                $ConfigObject.ConfigItem = "No of policies : $AllGlobalPolicy "
                $ConfigObject.ConfigData = "<B>No of Global policies defined is more than 5.</B>"
                $ConfigObject.InfoText = "You currently have $($AllGlobalPolicy) Global policies defined. We have found 5 Global Policy to be optimal for most organizations. Please consider reducing your current global policies count."
                $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Fail")  
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)
            }
            elseif ($AllGlobalPolicy -gt 10 ) {                    
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>High Number of Global policies defined<B>"
                $ConfigObject.ConfigItem = "No of policies : $AllGlobalPolicy "
                $ConfigObject.ConfigData = "You currently have $($AllGlobalPolicy) Global policies defined. We have found 5 Global Policy to be optimal for most organizations"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")  
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)

            }
            if ($ScopedGlobalPolicy -gt 10 -and $ScopedGlobalPolicy -lt 30 ) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>High Number of Scoped policies defined<B>"
                $ConfigObject.InfoText = "You currently have $($ScopedGlobalPolicy) Scoped policies defined. We have found 10 Scoped Policy to be optimal for most organizations. Please consider reducing your current scoped policies count."
                $ConfigObject.ConfigItem = "No of policies : $ScopedGlobalPolicy "
                $ConfigObject.ConfigData = "<B>No of parent Labels defined is more than 10.</B>"
                $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Fail")   
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)
            }
            elseif ($ScopedGlobalPolicy -gt 29 ) {                    
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>High Number of Scoped policies defined<B>"
                $ConfigObject.ConfigItem = "No of policies : $ScopedGlobalPolicy "
                $ConfigObject.ConfigData = "You currently have $($ScopedGlobalPolicy) Scoped policies defined. We have found 10 Scoped Policy to be optimal for most organizations."
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")  
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)

            }
            if ($ParentLabels.count -gt 5) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>High number of parent labels defined<B>"
                $ParentLabelsString = $null
                foreach ($ParentLabelkey in $ParentLabels.Keys) {
                    if ($null -ne $ParentLabelsString) {
                        $ParentLabelsString += ", $ParentLabelkey"
                    }
                    else {
                        $ParentLabelsString += "$ParentLabelkey"
                    }
                }
                $ConfigObject.ConfigItem = "Parent Labels : $($ParentLabelsString)"
                $ConfigObject.ConfigData = "<B>No. of parent labels defined is more than 5.</B>"
                $ConfigObject.InfoText = "You currently have $($ParentLabels.count) parent labels defined. We have found 5 parent labels to be optimal for most organizations. Please consider reducing your current label count. "
                $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Pass")   
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)

            }
            elseif ($ParentLabels.count -gt 10 ) {                    
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "High number of parent labels defined"
                $ParentLabelsString = $null
                foreach ($ParentLabelkey in $ParentLabels.Keys) {
                    if ($null -ne $ParentLabelsString) {
                        $ParentLabelsString += ", $ParentLabelkey"
                    }
                    else {
                        $ParentLabelsString += "$ParentLabelkey"
                    }
                }
                $ConfigObject.ConfigItem = "<B>Parent Labels</B> : $($ParentLabelsString)"
                $ConfigObject.ConfigData = "You currently have $($ParentLabels.count) parent labels defined. We have found 5 parent labels to be optimal for most organizations."
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")  
                $ConfigObjectList += $ConfigObject 
                $this.AddConfig($ConfigObject)

            }

            # For >5 sublabel count
        
            if ($($($ParentSubLabelAssociation.Keys).count) -gt 0) {
                $ParentLabelsWithHighSublabelCountString = $null
                foreach ($ParentLabelID in $($ParentSubLabelAssociation.Keys)) {
                    $AllSubLabels = $ParentSubLabelAssociation[$ParentLabelID] #all sublabels within a parent label
                
                    if ($($AllSubLabels.count) -gt 5) {
                        if ($null -ne $ParentLabelsWithHighSublabelCountString) {
                            $ParentLabelsWithHighSublabelCountString += ", $($ParentNameForSubLabelAssociation[$ParentLabelID])"
                        }
                        else {
                            $ParentLabelsWithHighSublabelCountString += "$($ParentNameForSubLabelAssociation[$ParentLabelID])"
                        }
                    }

                }
                if ($null -ne $ParentLabelsWithHighSublabelCountString) {
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.Object = "<B>High number of sub-labels defined</B>"
                    $ConfigObject.ConfigItem = "Parent Labels : $($ParentLabelsWithHighSublabelCountString)"
                    $ConfigObject.ConfigData = "<B>No. of sub-labels defined is more than 5.</B>"
                    $ConfigObject.InfoText = "You currently have more than 5 sub-labels defined for 1 or more parent labels. We have found 5 sub-labels to be optimal for most organizations. Please consider reducing your current sub-label count."
                    $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Pass") 
                    $ConfigObjectList += $ConfigObject 
                    $this.AddConfig($ConfigObject)
                }
            }

        
            If ($LabelPolicyExists -eq $False) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "No Active Policy defined"
                $ConfigObject.ConfigItem = "No Active Policy defined"
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