using module "..\MCCA.psm1"

class IRM101 : MCCACheck {
    <#
    

    #>

    IRM101() {
        $this.Control = "IRM-101"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Insider Risk Management"
        $this.Name = "Create customized or use default insider risk management policies for offensive language"
        $this.PassText = "Your organization has set up IRM policies for offensive language"
        $this.FailRecommendation = "Your organization should set up IRM policies for offensive language"
        $this.Importance = "Microsoft recommends that your organization create an insider risk management policy to detect, investigate, and take action on offensive and abusive behavior. Detecting and taking action to prevent offensive and abusive behavior is a critical component of preventing risk."
        $this.ExpandResults = $True
        $this.ItemName = "Policy"
        $this.DataType = "User Groups"
        $this.Links = @{
            "Getting started with Insider risk management" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-configure?view=o365-worldwide"
            "Compliance Center - Insider Risk Management" = "https://compliance.microsoft.com/insiderriskmgmt"
            "Insider risk management policies" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-policies"
        }
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        if ($Config["GetInsiderRiskPolicy"] -eq "Error") {
            $this.Completed = $false
        }
        else {
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
                . $UtilityFile.FullName
            }
            
            $Template = "WorkplaceThreat"
            $LogFile = $this.LogFile

            
            $ConfigObjectList = Get-IRMConfigurationPolicy -Config $Config -Templates @($Template) -LogFile $LogFile
            Foreach ($ConfigObject in $ConfigObjectList) {
                $this.AddConfig($ConfigObject)
            }
            

            $hasRemediation = $this.Config | Where-Object { $_.RemediationAction -ne '' }
            if ($($hasremediation.count) -gt 0) {
                $this.MCCARemediationInfo = New-Object -TypeName MCCARemediationInfo -Property @{
                    RemediationAvailable = $True
                    RemediationText      = "You need to connect to Exchange Online Center PowerShell to execute the below commands. Please follow steps defined in <a href = 'https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'> Connect to Exchange Online Center PowerShell</a>."
                }
            }
            $this.Completed = $True
        }
        
    }

}