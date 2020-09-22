using module "..\MCCA.psm1"

class Audit102 : MCCACheck {
    <#
    this is to valide if tenant has high serverity alert policies or not

    #>

    Audit102() {
        $this.Control = "Audit-102"
        $this.ParentArea = "Discovery & Response"
        $this.Area = "Audit"
        $this.Name = "Configure Alert Policies"
        $this.PassText = "Your organization has configured alert policies"
        $this.FailRecommendation = "Your organization should configure alert policies"
        $this.Importance = "Your organization should configure alert policies to send notifications on activities that are indicators of a potential security issue or data breach. Office 365 provides built-in alert policies that are turned on by default."
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ExpandResults = $True
        $this.ObjectType = "Alert Policy"
        $this.ItemName = "Severity"
        $this.DataType = "Email notifications"
        $this.Links = @{
            "Turn on audit log search" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/turn-audit-log-search-on-or-off?view=o365-worldwide"
            "Security & Compliance Console : Alert Policies" = "https://protection.office.com/?rfr=CMv3#/alertpolicies"
            "Learn more about alert policies" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/alert-policies?redirectSourcePath=%252farticle%252f8927b8b9-c5bc-45a8-a9f9-96c732e58264&view=o365-worldwide"
            "Compliance Manager - Audit Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Audit%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22OutOfScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {
        if ($Config["GetProtectionAlert"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $ConfigObjectList = @()
            $PoliciesExist = $false
            ForEach ($AlertPolicy in $Config["GetProtectionAlert"]) { 

                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "$($AlertPolicy.Name)"
                $ConfigObject.ConfigItem = "$($AlertPolicy.Severity)" 
                if($($AlertPolicy.Disabled) -eq $false)
                {
                    $PoliciesExist = $True
                    if($($AlertPolicy.NotificationEnabled) -eq $True)
                    {
                        $ConfigObject.ConfigData = $($AlertPolicy.NotifyUser)
                        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")
                        $this.AddConfig($ConfigObject)

                    }else{
                        $ConfigObject.ConfigData = "Email notifications not enabled"
                        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")
                        $this.AddConfig($ConfigObject)

                    }

                }else{
                    $ConfigObject.ConfigData = "Alert policy not enabled"
                    $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")
                    $this.AddConfig($ConfigObject)


                }
            }
            If ($PoliciesExist -eq $False) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "No active high severity policies were found"
                $ConfigObject.ConfigItem = "No active high severity policies"
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