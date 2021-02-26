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
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Turn on audit log search" = "https://aka.ms/mcca-aa-docs-action-turn-on"
                "Security & Compliance Console : Alert Policies" = "https://aka.ms/mcca-gcch-aa-2-compliance-center"
                "Learn more about alert policies" = "https://aka.ms/mcca-aa-docs-learn-more"
                "Compliance Manager - Audit Actions" = "https://aka.ms/mcca-gcch-aa-compliance-manager"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Turn on audit log search" = "https://aka.ms/mcca-aa-docs-action-turn-on"
                "Security & Compliance Console : Alert Policies" = "https://aka.ms/mcca-dod-aa-2-compliance-center"
                "Learn more about alert policies" = "https://aka.ms/mcca-aa-docs-learn-more"
                "Compliance Manager - Audit Actions" = "https://aka.ms/mcca-dod-aa-compliance-manager"
            }
        }else
        {
        $this.Links = @{
            "Turn on audit log search" = "https://aka.ms/mcca-aa-docs-action-turn-on"
            "Security & Compliance Console : Alert Policies" = "https://aka.ms/mcca-aa-2-compliance-center"
            "Learn more about alert policies" = "https://aka.ms/mcca-aa-docs-learn-more"
            "Compliance Manager - Audit Actions" = "https://aka.ms/mcca-aa-compliance-manager"
        }
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