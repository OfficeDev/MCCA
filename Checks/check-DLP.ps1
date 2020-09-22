using module "..\MCCA.psm1"

class DLP : MCCACheck {
    <#
    
       
    #>
    $SIT = $null
    $RemediationPolicyName = $null

    DLP($InfoParams) {
        $this.Control = $InfoParams["Control"]
        $this.ParentArea = $InfoParams["ParentArea"]
        $this.Area = $InfoParams["Area"]
        $this.Name = $InfoParams["Name"]
        $this.RemediationPolicyName = $InfoParams["RemediationPolicyName"]
        $this.PassText = $InfoParams["PassText"]
        $this.FailRecommendation = $InfoParams["FailRecommendation"]
        $this.Importance = $InfoParams["Importance"]
        $this.CheckType = [CheckType]::ObjectPropertyValue
        $this.ObjectType = "DLP Policy"
        $this.ItemName = "Sensitive Information Type"
        $this.DataType = "Remarks"
        $this.SIT = $InfoParams["SIT"]
        $this.Links = $InfoParams["Links"]
    
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {
        if (($Config["GetDlpComplianceRule"] -eq "Error") -or ($Config["GetDlpCompliancePolicy"] -eq "Error")) {
            $this.Completed = $false
        }
        else {
            if(($null -eq $($this.SIT)) -or ($($this.SIT) -eq ""))
            {
                $this.ExpandResults = $false
                $CheckNameDisplay = $this.RemediationPolicyName
                $CheckNameDisplayString = $CheckNameDisplay.Substring(5)     
                $this.Importance += "<div><span style='color:#cc9900;'>Note&nbsp;:&nbsp;</span>We currently do not support SITs for DLP policies for $CheckNameDisplayString for the geolocations for which this report is generated. Please review your DLP policies to ensure you are protected.</div>"
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Pass")   
                $this.AddConfig($ConfigObject)
                

            }
            else
            {
            $SensitiveTypes = @{}
            foreach ($SIT in $this.SIT) {
                $SensitiveTypes[$SIT] = $null
            }
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
                . $UtilityFile.FullName
            }
            $Name = "$($this.RemediationPolicyName)"
            if ($Name.length -gt 60) { $Name = $Name.substring(0, 60) }

            $LogFile = $this.LogFile
            $ConfigObjectList = Get-DLPPolicyValidation -SensitiveTypes $SensitiveTypes -Config $Config -Name $Name -LogFile $LogFile
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
            $this.ExpandResults = $True
        }
        $this.Completed = $True
    }
        
    }

}