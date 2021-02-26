using module "..\MCCA.psm1"

class IRM102 : MCCACheck {
    <#
    

    #>

    IRM102() {
        $this.Control = "IRM-102"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Insider Risk Management"
        $this.Name = "Create customized or use default insider risk management policies for departing employee data theft"
        $this.PassText = "Your organization has set up IRM policies for departing employee data theft"
        $this.FailRecommendation = "Your organization should set up IRM policies for departing employee data theft"
        $this.Importance = "Your organization should create an insider risk management policy to detect, investigate, and take action on departing employee data theft. Insider risk management in Microsoft 365 leverages an HR connector and selected indicators to alert you of any user activity related to data theft among departing employees."
        $this.ExpandResults = $True
        $this.ItemName = "Policy"
        $this.DataType = "User Groups"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Getting started with Insider risk management" = "https://aka.ms/mcca-irm-docs-action"
                "Compliance Center - Insider Risk Management" = "https://aka.ms/mcca-gcch-irm-compliance-center"
                "Insider risk management policies" = "https://aka.ms/mcca-irm-docs-learn-more"
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Getting started with Insider risk management" = "https://aka.ms/mcca-irm-docs-action"
                "Compliance Center - Insider Risk Management" = "https://aka.ms/mcca-dod-irm-compliance-center"
                "Insider risk management policies" = "https://aka.ms/mcca-irm-docs-learn-more"
            }  
        }else
        {
        $this.Links = @{
            "Getting started with Insider risk management" = "https://aka.ms/mcca-irm-docs-action"
            "Compliance Center - Insider Risk Management" = "https://aka.ms/mcca-irm-compliance-center"
            "Insider risk management policies" = "https://aka.ms/mcca-irm-docs-learn-more"
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
            
            $Template = "IntellectualPropertyTheft"
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