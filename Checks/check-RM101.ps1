using module "..\MCCA.psm1"

class RM101 : MCCACheck {
    <#
    

    #>

    RM101() {
        $this.Control = "RM-101"
        $this.ParentArea = "Microsoft Information Governance"
        $this.Area = "Records Management"
        $this.Name = "Declare Data as Records by Creating & Publishing a Record Label"
        $this.PassText = "Your organization is using record labels to declare data as records"
        $this.FailRecommendation = "Your organization should use record labels to declare data as records"
        $this.Importance = "Your organization should use records management to manage regulatory, legal, and business-critical records across corporate data. By using retention labels to declare records, you can implement a single, consistent records-management strategy across all of Office 365."
        $this.ExpandResults = $True
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ObjectType = "Policy Name"
        $this.ItemName = "Labels"
        $this.DataType = "Remarks"
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Overview of Records"              = "https://aka.ms/mcca-rm-docs-records"
                "Compliance Center - Records Management"                       = "https://aka.ms/mcca-gcch-rm-compliance-center"
                "Records management in Microsoft 365" = "https://aka.ms/mcca-rm-docs-records-management"
                "Compliance Manager - RM Actions" = "https://aka.ms/mcca-gcch-rm-compliance-manager"
            }   
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Overview of Records"              = "https://aka.ms/mcca-rm-docs-records"
                "Compliance Center - Records Management"                       = "https://aka.ms/mcca-dod-rm-compliance-center"
                "Records management in Microsoft 365" = "https://aka.ms/mcca-rm-docs-records-management"
                "Compliance Manager - RM Actions" = "https://aka.ms/mcca-dod-rm-compliance-manager"
            }  
        }else
        {
        $this.Links = @{
            "Overview of Records"              = "https://aka.ms/mcca-rm-docs-records"
            "Compliance Center - Records Management"                       = "https://aka.ms/mcca-rm-compliance-center"
            "Records management in Microsoft 365" = "https://aka.ms/mcca-rm-docs-records-management"
            "Compliance Manager - RM Actions" = "https://aka.ms/mcca-rm-compliance-manager"
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {               
        if (($Config["GetRetentionComplianceRule"] -eq "Error") -or ($Config["GetRetentionCompliancePolicy"] -eq "Error") -or ($Config["GetComplianceTag"] -eq "Error")) {
            $this.Completed = $false
        }
        else {
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
        
                . $UtilityFile.FullName
                
            }
            
            $LogFile = $this.LogFile
            $Mode= "Publish"
            $ConfigObjectList = Get-RMPolicyValidation -LogFile $LogFile -Mode $Mode

            Foreach ($ConfigObject in $ConfigObjectList) {
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

