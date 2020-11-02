using module "..\MCCA.psm1"

class CC101 : MCCACheck {
    <#
    

    #>

    CC101() {
        $this.Control = "CC-101"
        $this.ParentArea = "Insider Risk"
        $this.Area = "Communication Compliance"
        $this.Name = "Enable Communication Compliance in O365"
        $this.PassText = "Your organization has enabled Communication Compliance in O365"
        $this.FailRecommendation = "Your organization should enable Communication Compliance in O365"
        $this.Importance = "Your organization should use communication compliance to scan internal and external communications for policy matches so they can be examined by designated reviewers."
        $this.ExpandResults = $True
        $this.ItemName = "Role"
        $this.DataType = "Role Groups </B> ( Having 1 or more members)"
        $this.Links = @{
            "Communication compliance in Microsoft 365"     = "https://go.microsoft.com/fwlink/?linkid=2107258"
            "Compliance Center - Communication Compliance" = "https://compliance.microsoft.com/supervisoryreview"
            "Compliance Manager - CC Actions" = "https://compliance.microsoft.com/compliancescore?filter=%7B%22Solution%22:%5B%22Communication%20compliance%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
        }
    
    }

    <#
    
        RESULTS CC Admin, CC Analyst, CC Investigator and CC Viewer
    #>

    GetResults($Config) {   

        try {
            $SreviewAdminRoleGroups = Get-RoleGroup -ErrorAction:SilentlyContinue | Where-Object { $_.Roles -Like "*Supervisory Review Administrator*" -and $null -ne $_.Members }  
            $CaseManagementRoleGroups = Get-RoleGroup -ErrorAction:SilentlyContinue | Where-Object { $_.Roles -Like "*Case Management*" -and $null -ne $_.Members }  
            $ComplianceAdministratorRoleGroups = Get-RoleGroup -ErrorAction:SilentlyContinue | Where-Object { $_.Roles -Like "*Compliance Administrator*" -and $null -ne $_.Members } 
        
        }
        catch {
            $SreviewAdminRoleGroups = "Error"
            $CaseManagementRoleGroups = "Error"
            $ComplianceAdministratorRoleGroups = "Error"
        }
        if (($SreviewAdminRoleGroups -eq "Error") -or ($CaseManagementRoleGroups -eq "Error") -or ($ComplianceAdministratorRoleGroups -eq "Error")) {
            $this.Completed = $false
        }
        else {
            $UtilityFiles = Get-ChildItem "$PSScriptRoot\..\Utilities"

            ForEach ($UtilityFile in $UtilityFiles) {
                . $UtilityFile.FullName
            } 
            $LogFile = $this.LogFile

            $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $SreviewAdminRoleGroups -Role "Supervisory Review Administrator" -LogFile $LogFile
            Foreach ($ConfigObject in $ConfigObjectList) {
                $this.AddConfig($ConfigObject)
            }
            $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $CaseManagementRoleGroups -Role "Case Management" -LogFile $LogFile
            Foreach ($ConfigObject in $ConfigObjectList) {
                $this.AddConfig($ConfigObject)
            }
            $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $ComplianceAdministratorRoleGroups -Role "Compliance Administrator" -LogFile $LogFile
            Foreach ($ConfigObject in $ConfigObjectList) {
                $this.AddConfig($ConfigObject)
            }
            # New roles post CC july release
            #$CCAdminRoleGroups = Get-RoleGroup | Where-Object {$_.Roles  -Like "*Communication Compliance Admin*" -and $_.Members -ne $null} 
            #$CCAnalystRoleGroups = Get-RoleGroup | Where-Object {$_.Roles  -Like "*Communication Compliance Analyst*" -and $_.Members -ne $null}  
            #$CCInvesRoleGroups = Get-RoleGroup | Where-Object {$_.Roles  -Like "*Communication Compliance Investigator*" -and $_.Members -ne $null}  
            #$CCViewRoleGroups = Get-RoleGroup | Where-Object {$_.Roles  -Like "*Communication Compliance Viewer*" -and $_.Members -ne $null}  
       
            <#
        $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $CCAdminRoleGroups -Role "Communication Compliance Admin"
        Foreach ($ConfigObject in $ConfigObjectList)
        {
            $this.AddConfig($ConfigObject)
        }
         $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $CCAnalystRoleGroups -Role "Communication Compliance Analyst"
        Foreach ($ConfigObject in $ConfigObjectList)
        {
            $this.AddConfig($ConfigObject)
        }
         $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $CCInvesRoleGroups -Role "Communication Compliance Investigator"
        Foreach ($ConfigObject in $ConfigObjectList)
        {
            $this.AddConfig($ConfigObject)
        }
         $ConfigObjectList = Get-RoleGroupwithMembers -RoleGroups $CCViewRoleGroups -Role "Communication Compliance Viewer"
        Foreach ($ConfigObject in $ConfigObjectList)
        {
            $this.AddConfig($ConfigObject)
        }#>

            $hasRemediation = $this.Config | Where-Object { $_.RemediationAction -ne '' }
            if ($($hasremediation.count) -gt 0) {
                $this.MCCARemediationInfo = New-Object -TypeName MCCARemediationInfo -Property @{
                    RemediationAvailable = $True
                    RemediationText      = "You need to connect to Security & Compliance Center PowerShell to execute the below commands. Please follow steps defined in <a href = 'https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps'> Connect to Security & Compliance Center PowerShell</a>."
                }
            }
            $this.Completed = $True
        }
      
    }

}