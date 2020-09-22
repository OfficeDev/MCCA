using module "..\MCCA.psm1"

class IG102 : MCCACheck {
    <#
    

    #>

    IG102() {
        $this.Control = "IG-102"
        $this.ParentArea = "Microsoft Information Governance"
        $this.Area = "Information Governance"
        $this.Name = "Use Data Retention Labels and Policies"
        $this.PassText = "Your organization is using retention policies by publishing a retention label"
        $this.FailRecommendation = "Your organization should use retention policies by publishing a retention label"
        $this.Importance = "Your organization should apply retention labels to content when it matches specific conditions (such as containing specific keywords or types of sensitive information)."
        $this.ExpandResults = $True
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ObjectType = "Retention Policies"
        $this.ItemName = "Labels"
        $this.DataType = "Remarks"
        $this.Links = @{
            "Learn More Overview of retention labels"     = "https://docs.microsoft.com/en-us/microsoft-365/compliance/labels?redirectSourcePath=%252farticle%252faf398293-c69d-465e-a249-d74561552d30&view=o365-worldwide"
            "Overview of retention policies"              = "https://docs.microsoft.com/en-us/microsoft-365/compliance/retention-policies?view=o365-worldwide"
            "Compliance Center - Information Governance" = "https://compliance.microsoft.com/informationgovernance?"
            "Compliance Manager - IG Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Information%20governance%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22OutOfscope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS CC Admin, CC Analyst, CC Investigator and CC Viewer
    #>

    GetResults($Config) {               
        if (($Config["GetRetentionComplianceRule"] -eq "Error") -or ($Config["GetRetentionCompliancePolicy"] -eq "Error")) {
            $this.Completed = $false
        }
        else {
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
        
                . $UtilityFile.FullName
                
            }
            
            $LogFile = $this.LogFile
            $Mode = "Publish"
            $ConfigObjectList = Get-RetentionPolicyValidation -LogFile $LogFile -Mode $Mode

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