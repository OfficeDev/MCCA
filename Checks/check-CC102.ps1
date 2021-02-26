using module "..\MCCA.psm1"

class CC102 : MCCACheck {
    <#
    

    #>

    CC102() {
        $this.Control = "CC-102"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Communication Compliance"
        $this.Name = "Monitor Communications for Offensive or Threatening Language"
        $this.PassText = "Your organization has defined policies to monitor internal communications"
        $this.FailRecommendation = "Your organization should define policies to monitor internal communications"
        $this.Importance = "Your organization should use communication compliance to monitor internal communication for offensive and threatening language. You can create a policy that uses pretrained classifier to detect content containing profanities or language that might be considered threatening or harrassment."
        $this.ExpandResults = $True
        $this.ItemName = "Policy"
        $this.DataType = "Policy Status"
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
        if ($Config["GetSupervisoryReviewPolicyV2"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()  
            $PolicyExists = $False
            ForEach ($Policy in $Config["GetSupervisoryReviewPolicyV2"]) {
                $PolicyExists = $True
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.ConfigItem = "$($Policy.Name)"
                $ConfigObject.ConfigData = $($Policy.PolicyStatus)
                if ($($Policy.PolicyStatus) -ieq "Active") {
                    $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")  
                }
                else {
                    $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            

                }
                $ConfigObjectList += $ConfigObject
                $this.AddConfig($ConfigObject)

            }
        
            If ($PolicyExists -eq $False) {
                $ConfigObject = [MCCACheckConfig]::new()

                $ConfigObject.ConfigItem = "No Active Policy Defined"
                $ConfigObject.ConfigData = "No Active Policy Defined"
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