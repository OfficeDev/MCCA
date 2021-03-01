using module "..\MCCA.psm1"

class CC103 : MCCACheck {
    <#
    

    #>

    CC103() {
        $this.Control = "CC-103"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Communication Compliance"
        $this.Name = "Remediate Corporate Policy Violation"
        $this.PassText = "Your organization currently has no corporate policy violations"
        $this.FailRecommendation = "Your organization needs to remediate corporate policy violations"
        $this.Importance = "Your organization should use communication compliance to scan internal and external communications for policy matches so they can be examined by designated reviewers. Reviewers can investigate scanned communications and take appropriate remediation actions."
        $this.ExpandResults = $True
        $this.ItemName = "Communication Compliance Remediation"
        $this.DataType = "Items pending Review"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Communication compliance in Microsoft 365"     = "https://aka.ms/mcca-cc-docs-learn-more"
                "Compliance Center - Communication Compliance" = "https://aka.ms/mcca-gcch-cc-compliance-center"
                "Compliance Manager - CC Actions" = "https://aka.ms/mcca-gcch-cc-compliance-manager"
              }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Communication compliance in Microsoft 365"     = "https://aka.ms/mcca-cc-docs-learn-more"
                "Compliance Center - Communication Compliance" = "https://aka.ms/mcca-dod-cc-compliance-center"
                "Compliance Manager - CC Actions" = "https://aka.ms/mcca-dod-cc-compliance-manager"
         }
        }else
        {
        $this.Links = @{
            "Communication compliance in Microsoft 365"     = "https://aka.ms/mcca-cc-docs-learn-more"
            "Compliance Center - Communication Compliance" = "https://aka.ms/mcca-cc-compliance-center"
            "Compliance Manager - CC Actions" = "https://aka.ms/mcca-cc-compliance-manager"

        }
        }
    }

    <#
    
        RESULTS CC Admin, CC Analyst, CC Investigator and CC Viewer
    #>

    GetResults($Config) {         
        if (($Config["GetSupervisoryReviewOverallProgressReport"] -eq "Error") -or ($Config["GetSupervisoryReviewPolicyV2"] -eq "Error")) {
            $this.Completed = $false
        }     
        else {
            $ConfigObjectList = @() 
            $SupervisoryReviewOverallProgressReport = $Config["GetSupervisoryReviewOverallProgressReport"]
            if ( $null -eq $SupervisoryReviewOverallProgressReport) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.ConfigItem = "Communication Compliance Policy Matches"

                $supervisory = $Config["GetSupervisoryReviewPolicyV2"]

                if ($($supervisory.count) -eq 0) {
                    $ConfigObject.ConfigData = "No communication Policy defined"
                }
                else {
                    $ConfigObject.ConfigData = "User does not have access to policy review"
                }
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")    
                $ConfigObjectList += $ConfigObject      
                $this.AddConfig($ConfigObject)

            }
            elseif ($($SupervisoryReviewOverallProgressReport.Pending) -eq 0) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.ConfigItem = "Communication Compliance Policy Matches"
                $ConfigObject.ConfigData = "None"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")  
                $ConfigObjectList += $ConfigObject
                $this.AddConfig($ConfigObject)

            }
        
            else {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.ConfigItem = "Communication Compliance Policy Matches"
                $ConfigObject.ConfigData = "$($SupervisoryReviewOverallProgressReport.Pending)"
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