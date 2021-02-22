using module "..\MCCA.psm1"

class IRM103 : MCCACheck {
    <#
    

    #>

    IRM103() {
        $this.Control = "IRM-103"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Insider Risk Management"
        $this.Name = "Create customized or use default insider risk management policies for data leaks"
        $this.PassText = "Your organization has set up IRM policies for data leaks"
        $this.FailRecommendation = "Your organization should set up IRM policies for data leaks"
        $this.Importance = "Microsoft recommends that your organization create an insider risk management policy to detect, investigate, and take action on data leaks. Data leaks can include accidental oversharing of information outside your organization or data theft with malicious intent."
        $this.ExpandResults = $True
        $this.ItemName = "Policy"
        $this.DataType = "User Groups"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
            
                "Getting started with Insider risk management" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-configure?view=o365-worldwide"
                "Compliance Center - Insider Risk Management" = "https://compliance.microsoft.us/insiderriskmgmt"
                "Insider risk management policies" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-policies"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
            
                "Getting started with Insider risk management" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-configure?view=o365-worldwide"
                "Compliance Center - Insider Risk Management" = "https://compliance.apps.mil/insiderriskmgmt"
                "Insider risk management policies" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-policies"
            }
        }else
        {
        $this.Links = @{
            
            "Getting started with Insider risk management" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-configure?view=o365-worldwide"
            "Compliance Center - Insider Risk Management" = "https://compliance.microsoft.com/insiderriskmgmt"
            "Insider risk management policies" = "https://docs.microsoft.com/microsoft-365/compliance/insider-risk-management-policies"
        }
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
            #LeakOfInformation OR DisgruntledEmployeeDataLeak OR HighValueEmployeeDataLeak
            $Template = @("LeakOfInformation","DisgruntledEmployeeDataLeak","HighValueEmployeeDataLeak")
            $LogFile = $this.LogFile

            
            $ConfigObjectList = Get-IRMConfigurationPolicy -Config $Config -Templates $Template -LogFile $LogFile
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