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