
using module "..\MCCA.psm1"
$ExchangeString = "Exchange"
$SharePointString = "SharePoint"
$OneDriveString = "OneDrive"
$TeamsString = "Teams"
$DevicesString = "Devices"

Function Get-DLPPolicyValidation {
    Param
    (
        $SensitiveTypes,
        $Config,
        $LogFile,
        $Name
    )
    $ConfigObjectList = @()
    try {
        $SensitiveTypesWorkloadMapping = @()
        foreach ($SIT in $SensitiveTypes.keys) {
                
            $SensitiveTypesWorkloadMapping += New-Object -TypeName PSObject @{
                Name              = $SIT
                $ExchangeString   = $false
                $SharePointString = $false
                $OneDriveString   = $false
                $TeamsString      = $false
                $DevicesString    = $false
            
            }
        }
        #Getting Custom SIT
        if($($Config["GetDLPCustomSIT"]) -ne "Error")
        {
            $CustomSIT = $($Config["GetDLPCustomSIT"]).Name
            $CustomSITHashTable = @{}
            foreach($SIT in $CustomSIT)
            {
                $CustomSITHashTable[$SIT] = $null
            }

        }

        ForEach ($CompliancePolicy in $Config["GetDlpCompliancePolicy"]) {   
            $PolicySensitiveType = New-Object System.Collections.Generic.HashSet[String]          
            $PolicySensitiveType = Get-PolicySensitiveType -Config $Config -CompliancePolicy $CompliancePolicy -SensitiveTypes $SensitiveTypes
            if($($Config["GetDLPCustomSIT"]) -ne "Error")
            {
                $CustomSensitiveType = Get-PolicySensitiveType -Config $Config -CompliancePolicy $CompliancePolicy -SensitiveTypes $CustomSITHashTable
                $CustomSensitiveTypeText = $null
                foreach ($SIT in $CustomSensitiveType) {
                    if ($null -ne $CustomSensitiveTypeText) {
                        $CustomSensitiveTypeText += ", $SIT"
                    }
                    else {
                        $CustomSensitiveTypeText += "$SIT"
                    }
                }
            }
            if (($CompliancePolicy.Mode -ieq "enable") ) {
                $WorkloadsStatus = Get-AllLocationenabled -CompliancePolicy $CompliancePolicy 
                $EnabledWorkload = $null
                $DisabledWorkload = ""       
                $PolicySensitiveTypeText = $null
                foreach ($Workload in ($WorkloadsStatus.Keys | Sort-Object -CaseSensitive)) {
                    if ($WorkloadsStatus[$Workload] -eq $true) {
                        if ( $null -ne $EnabledWorkload) {
                            $EnabledWorkload += ", $($Workload)"
                        }
                        else {
                            $EnabledWorkload += "$($Workload)"
                        }
                        foreach ($SIT in $PolicySensitiveType) {
                            if ($SITToChange = $SensitiveTypesWorkloadMapping | Where-Object { $_.Name -eq $SIT }) {
                                $SITToChange.$($Workload) = $true
                            }                        
                        }
                                     
                    }
                    else {
                        $DisabledWorkload += "$($Workload) "                 
                    }           
                }
                
                foreach ($SIT in $PolicySensitiveType) {
                    if ($null -ne $PolicySensitiveTypeText) {
                        $PolicySensitiveTypeText += ", $SIT"
                    }
                    else {
                        $PolicySensitiveTypeText += "$SIT"
                    }
                }
            
               
                If ($PolicySensitiveType.Count -ne 0 ) {   
                    $ConfigObject = [MCCACheckConfig]::new()
                    $Workload = $CompliancePolicy.Workload
                    $ConfigObject.Object = "$($CompliancePolicy.Name)"
                    if($null -eq $CustomSensitiveTypeText)
                    {
                        $ConfigObject.ConfigItem = "$PolicySensitiveTypeText"
                    
                    }
                    else
                    {
                        $ConfigObject.ConfigItem = "$PolicySensitiveTypeText<br><strong>Custom SIT</strong> : $CustomSensitiveTypeText"
                    }
                    $ConfigData = ""
                    $ConfigObjectResult = @{}
                    $ConfigObjectResult = Set-ExchangeNotAllLocationEnabledConfigObject -ConfigObjectResult $ConfigObjectResult -CompliancePolicy $CompliancePolicy
                    $ConfigObjectResult = Set-SharePointNotAllLocationEnabledConfigObject -ConfigObjectResult $ConfigObjectResult -CompliancePolicy $CompliancePolicy
                    $ConfigObjectResult = Set-OneDriveNotAllLocationEnabledConfigObject  -ConfigObjectResult $ConfigObjectResult -CompliancePolicy $CompliancePolicy
                    $ConfigObjectResult = Set-TeamsNotAllLocationEnabledConfigObject  -ConfigObjectResult $ConfigObjectResult -CompliancePolicy $CompliancePolicy
                    $ConfigObjectResult = Set-DevicesNotAllLocationEnabledConfigObject  -ConfigObjectResult $ConfigObjectResult -CompliancePolicy $CompliancePolicy
                    $ConfigData = "<strong>Enabled Workloads </strong>: $($EnabledWorkload)<BR/>"
                    foreach ($ConfigResult in $ConfigObjectResult.keys) {
                        $ConfigData += "<strong>$ConfigResult </strong>: $($ConfigObjectResult[$ConfigResult])<BR/>"
                            
                    }
                    $NotInOrganizationAccessScope = $Config["GetDlpComplianceRule"] | Where-Object {$_.AccessScope -eq "NotInOrganization" -and $_.ParentPolicyName -eq "$($CompliancePolicy.Name)"} 
                    if($null -ne $NotInOrganizationAccessScope)
                    {
                        $ConfigData += "<strong>Access Scope</strong>: For users outside organization<BR/>"
                       
                    }else{
                        $ConfigData += "<strong>Access Scope</strong>: For users inside organization<BR/>"
                    }
                    $ConfigObject.ConfigData = "$ConfigData"
                    $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")   
                    $ConfigObjectList += $ConfigObject
    
                }
                
            }   
            else {
                If ($PolicySensitiveType.Count -ne 0 ) {   
                    $ConfigObject = [MCCACheckConfig]::new()
                    $Workload = $CompliancePolicy.Workload
                    $ConfigObject.Object = "$($CompliancePolicy.Name)"
                    $PolicySensitiveTypeText = $null
                    foreach ($SIT in $PolicySensitiveType) {
                        if ($null -ne $PolicySensitiveTypeText) {
                            $PolicySensitiveTypeText += ", $SIT"
                        }
                        else {
                            $PolicySensitiveTypeText += "$SIT"
                        }
                    }
                    if($null -eq $CustomSensitiveTypeText)
                    {
                        $ConfigObject.ConfigItem = "$PolicySensitiveTypeText"
                    
                    }
                    else
                    {
                        $ConfigObject.ConfigItem = "$PolicySensitiveTypeText<br><strong>Custom SIT</strong> : $CustomSensitiveTypeText"
                    }
                    $Mode = $($CompliancePolicy.Mode)
                    if ( $Mode -eq "TestWithoutNotifications") {
                        $Mode = "test without notifications"
                    }
                    elseif ($Mode -eq "Disable") {
                        $Mode = "disabled"
                    }
                    elseif ( $Mode -eq "TestWithNotifications") {
                        $Mode = "test with notifications"
                    }
                   
                    $ConfigObject.ConfigData = "<B>Policy is in $Mode state.<B>"
                    $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")   
    
                    $ConfigObjectList += $ConfigObject
                }
            }
               
        }
        $ConfigObjectList = Get-SensitiveTypesNotEnabled -SensitiveTypesWorkloadMapping $SensitiveTypesWorkloadMapping -ConfigObjectList $ConfigObjectList -LogFile $LogFile
         
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue while running MCCA. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
        
    }
    Return $ConfigObjectList
}

Function Get-NoPolicyRemediationAction {
    [CmdletBinding()]
    Param
    (
        $Name,
        $PendingSensitiveTypes
      
    )
    $RemediationActionScript = ""

    $PendingSensitiveTypesList = $PendingSensitiveTypes.split(",") 
    $LowCountSenstiveinfodetails = ""
    $HighCountSenstiveinfodetails = ""

    foreach ($PendingSensitiveType in $PendingSensitiveTypesList) {
        $PendingSensitiveTypetrim = $PendingSensitiveType.trim() 
        if ( $LowCountSenstiveinfodetails -eq "") {
            $LowCountSenstiveinfodetails += "@{Name =" + [char]34
            $HighCountSenstiveinfodetails += "@{Name =" + [char]34

        }
        else {
            $LowCountSenstiveinfodetails += ",@{Name =" + [char]34
            $HighCountSenstiveinfodetails += ",@{Name =" + [char]34

        }
        $LowCountSenstiveinfodetails += $PendingSensitiveTypetrim + [char]34
        $LowCountSenstiveinfodetails += ";minCount = ""1"""
        $LowCountSenstiveinfodetails += ";maxCount = ""5""}"
        $HighCountSenstiveinfodetails += $PendingSensitiveTypetrim + [char]34
        $HighCountSenstiveinfodetails += ";minCount = ""6""}"
    }
                

    $NewPolicyTemplateData = Get-Content "$PSScriptRoot\..\Templates\NewDLPPolicyTemplate.txt"
    if ($null -eq $NewPolicyTemplateData -or $NewPolicyTemplateData -eq "") {
        Write-Host "$(Get-Date) Template file does not exist/is corrupt in $PSScriptRoot\..\Templates\NewDLPPolicyTemplate.txt. Remediation wont be generated" -ForegroundColor Orange           
    }
    else {
        $NewPolicyTemplateData = $NewPolicyTemplateData.Replace("<NewPolicyName>", "$Name")
        $NewPolicyTemplateData = $NewPolicyTemplateData.Replace("<HighSensitiveInfoDetails>", "$HighCountSenstiveinfodetails")
        $NewPolicyTemplateData = $NewPolicyTemplateData.Replace("<LowSensitiveInfoDetails>", "$LowCountSenstiveinfodetails")
        $LowRuleName = "Low Volume $Name"
        if ($LowRuleName.length -gt 60) { $LowRuleName = $LowRuleName.substring(0, 60) }
        $HighRuleName = "High Volume $Name"
        if ($HighRuleName.length -gt 60) { $HighRuleName = $HighRuleName.substring(0, 60) }
        $NewPolicyTemplateData = $NewPolicyTemplateData.Replace("<HighVolumeRuleName>", "$HighRuleName")
        $NewPolicyTemplateData = $NewPolicyTemplateData.Replace("<LowVolumeRuleName>", "$LowRuleName")
  
        $RemediationActionScript += $NewPolicyTemplateData
        Write-Host "$(Get-Date) Generating Remediation Action for $Name" -ForegroundColor Yellow 
    }
              
    Return $RemediationActionScript
}
Function Get-PolicySensitiveType {
    Param
    (
        $Config,
        $CompliancePolicy,
        $SensitiveTypes
    )
    $PolicySensitiveTypes = New-Object System.Collections.Generic.HashSet[String]          
    foreach ($ComplianceRule in $Config["GetDlpComplianceRule"]) {

        if ($ComplianceRule.Mode -ieq "enforce" -and $CompliancePolicy.name -eq $($ComplianceRule.ParentPolicyName) ) {
            $SensitiveInformationContent = $ComplianceRule.ContentContainsSensitiveInformation

            foreach ($SensitiveType in $($SensitiveTypes.keys)) {
                if ($SensitiveInformationContent.Values -contains $SensitiveType) {
                    if (!$PolicySensitiveTypes.Contains($SensitiveType)) {
                        $PolicySensitiveTypes.Add("$SensitiveType") |  Out-Null 

                    }

                }
                if ($($SensitiveInformationContent.keys) -contains "groups") {
                    foreach ($SensitiveInformationGroupList in $SensitiveInformationContent) {
                        $SensitiveInformationGroups = $SensitiveInformationGroupList["groups"]
                        foreach ($SensitiveInformationGroupDefined in $SensitiveInformationGroups) {
                            $SensitiveInformationGroupDefinedValues = $SensitiveInformationGroupDefined.Values 
                            foreach ($SensitiveInformationGroupValue in $SensitiveInformationGroupDefinedValues) {
                                foreach ($SensitiveInformationGroupVal in $SensitiveInformationGroupValue) {
                                    if ($SensitiveInformationGroupVal.Values -contains $SensitiveType ) {
                                        if (!$PolicySensitiveTypes.Contains($SensitiveType)) {
                                            $PolicySensitiveTypes.Add("$SensitiveType") |  Out-Null 

                                        }
                                    }
      
                                }
     
                            }
                        }
                    }
    
                }
            }
                            

                           
        }
    }

    Return $PolicySensitiveTypes
}

Function Get-SensitiveTypesNotEnabled {
    Param
    (
        $SensitiveTypesWorkloadMapping,
        $LogFile,
        $ConfigObjectList
    )   

     
    $PendingSensitiveType = $null
    $PartialCoveredSIT = $null
    $PartialCoveredWorkload = $null
    foreach ($SensitiveTypes in $SensitiveTypesWorkloadMapping) {
        if (($SensitiveTypes.$ExchangeString -eq $false ) -and ($SensitiveTypes.$SharePointString -eq $false ) -and 
            ($SensitiveTypes.$TeamsString -eq $false ) -and ($SensitiveTypes.$OneDriveString -eq $false ) -and ($SensitiveTypes.$DevicesString -eq $false ) ) {
            $PendingSensitiveType = Get-PartialSIT -PartialCoveredSIT $PendingSensitiveType -SensitiveTypesName $($SensitiveTypes.Name)

        }
        else {

            if ($SensitiveTypes.$ExchangeString -eq $false ) {
                $PartialCoveredSIT = Get-PartialSIT -PartialCoveredSIT $PartialCoveredSIT -SensitiveTypesName $($SensitiveTypes.Name)
                $PartialCoveredWorkload = Get-PartialSITWorkLoad -PartialCoveredWorkload $PartialCoveredWorkload -WorkloadName $ExchangeString
            }
    
            if ($SensitiveTypes.$SharePointString -eq $false ) {
                $PartialCoveredSIT = Get-PartialSIT -PartialCoveredSIT $PartialCoveredSIT -SensitiveTypesName $($SensitiveTypes.Name)
                $PartialCoveredWorkload = Get-PartialSITWorkLoad -PartialCoveredWorkload $PartialCoveredWorkload -WorkloadName $SharePointString
    
            } 
            if ($SensitiveTypes.$OneDriveString -eq $false ) {
                $PartialCoveredSIT = Get-PartialSIT -PartialCoveredSIT $PartialCoveredSIT -SensitiveTypesName $($SensitiveTypes.Name)
                $PartialCoveredWorkload = Get-PartialSITWorkLoad -PartialCoveredWorkload $PartialCoveredWorkload -WorkloadName $OneDriveString
            }
            if ($SensitiveTypes.$TeamsString -eq $false ) {
                $PartialCoveredSIT = Get-PartialSIT -PartialCoveredSIT $PartialCoveredSIT -SensitiveTypesName $($SensitiveTypes.Name)
                $PartialCoveredWorkload = Get-PartialSITWorkLoad -PartialCoveredWorkload $PartialCoveredWorkload -WorkloadName $TeamsString
            }
            if ($SensitiveTypes.$DevicesString -eq $false ) {
                $PartialCoveredSIT = Get-PartialSIT -PartialCoveredSIT $PartialCoveredSIT -SensitiveTypesName $($SensitiveTypes.Name)
                $PartialCoveredWorkload = Get-PartialSITWorkLoad -PartialCoveredWorkload $PartialCoveredWorkload -WorkloadName $DevicesString
            }
    
        }

        
     
        
    }
   

   
    if ($null -ne $PartialCoveredSIT) {
        $ConfigObject = [MCCACheckConfig]::new()
        $ConfigObject.Object = "<B>Policy defined but not protected on 1 or more workloads<B>"
        $ConfigObject.ConfigItem = "$PartialCoveredSIT"
        $ConfigObject.ConfigData = "<b>Affected Workloads</B> :  $PartialCoveredWorkload"
        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
        $ConfigObjectList += $ConfigObject
    }
    if ($null -ne $PendingSensitiveType) { 
                   
        $ConfigObject = [MCCACheckConfig]::new()
        $ConfigObject.Object = "<B>No active policy defined<B>"
        $ConfigObject.ConfigItem = "$PendingSensitiveType"
        $ConfigObject.ConfigData = "<b>Affected Workloads</B> :  $ExchangeString, $SharePointString, $TeamsString, $OneDriveString, $DevicesString"
        $ConfigObject.InfoText ="It is recommended that you set up DLP policies that block access for users external to your organization for all Sensitive Information Types on all workloads."
        try {
            $ConfigObject.RemediationAction = Get-NoPolicyRemediationAction -Name $Name -PendingSensitiveTypes $PendingSensitiveType -ErrorAction:SilentlyContinue        
        }
        catch {
            Write-Host "Warning:$(Get-Date) There was an issue in generating remediation script. Please review the script closely before running the same." -ForegroundColor:Yellow
            $ErrorMessage = $_.ToString()
            $StackTraceInfo = $_.ScriptStackTrace
            Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
        }
        $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
        $ConfigObjectList += $ConfigObject
    }
    
    Return $ConfigObjectList
}
function Get-PartialSIT {
    Param
    (
        $PartialCoveredSIT,
        $SensitiveTypesName

    )
    if ((@($PartialCoveredSIT) -like "*$SensitiveTypesName*").Count -le 0) { 
        if ($null -eq $PartialCoveredSIT) {
            $PartialCoveredSIT += "$SensitiveTypesName"
        
        }
        else {            
            $PartialCoveredSIT += ", $SensitiveTypesName" 
        } 
    }
    return $PartialCoveredSIT
}
function Get-PartialSITWorkLoad {
    Param
    (
        $PartialCoveredWorkload,
        $WorkloadName

    )
    if ((@($PartialCoveredWorkload) -like "*$WorkloadName*").Count -le 0) {
        if ($null -eq $PartialCoveredWorkload) {
            $PartialCoveredWorkload += $WorkloadName
            
        }
        else {            
            $PartialCoveredWorkload += ", $WorkloadName" 
        } 
    }
    return $PartialCoveredWorkload
}


Function Get-AllLocationenabled {
    Param
    (
        $CompliancePolicy
    )
    $ExchangeLocation = $CompliancePolicy.ExchangeLocation
    $SharePointLocation = $CompliancePolicy.SharePointLocation
    $OneDriveLocation = $CompliancePolicy.OneDriveLocation
    $TeamsLocation = $CompliancePolicy.TeamsLocation
    $EndpointDlpLocation = $CompliancePolicy.EndpointDlpLocation   
    $WorkloadsStatus = @{}
    $WorkloadsStatus[$ExchangeString] = $false
    $WorkloadsStatus[$SharePointString] = $false
    $WorkloadsStatus[$OneDriveString] = $false
    $WorkloadsStatus[$TeamsString] = $false
    $WorkloadsStatus[$DevicesString] = $false
    if ((@($ExchangeLocation) -like 'All').Count -gt 0) {
        $WorkloadsStatus[$ExchangeString] = $true
    }
    if ((@($SharePointLocation) -like 'All').Count -gt 0) {
        $WorkloadsStatus[$SharePointString] = $true
    }
    if ((@($OneDriveLocation) -like 'All').Count -gt 0) {
        $WorkloadsStatus[$OneDriveString] = $true
    }
    if ((@($TeamsLocation) -like 'All').Count -gt 0) {
        $WorkloadsStatus[$TeamsString] = $true
    }
    if ((@($EndpointDlpLocation) -like 'All').Count -gt 0) {
        $WorkloadsStatus[$DevicesString] = $true
    }

    Return $WorkloadsStatus

    
}


Function Set-ExchangeNotAllLocationEnabledConfigObject {
    Param
    (
        
        $ConfigObjectResult,
        $CompliancePolicy
    )

    $ExchangeLocation = $CompliancePolicy.ExchangeLocation
    $ExchangeSenderException = $CompliancePolicy.ExchangeSenderException
    $ExchangeSenderMemberOf = $CompliancePolicy.ExchangeSenderMemberOf
    $ExchangeSenderMemberOfException = $CompliancePolicy.ExchangeSenderMemberOfException

    if (((@($ExchangeLocation) -like 'All').Count -lt 1)) {          
        if (@($ExchangeLocation).count -ne 0) {
            
            $ConfigObjectResult["Included Exchange Groups"] += "$ExchangeLocation " 
                    
        }
    }

    if ($ExchangeSenderMemberOf.count -ne 0) {
    
        if ($ConfigObjectResult.contains("Included Exchange Groups")) {
            $ConfigObjectResult["Included Exchange Groups"] += ", $ExchangeSenderMemberOf " 
        }
        else {
            $ConfigObjectResult["Included Exchange Groups"] = "$ExchangeSenderMemberOf " 

        }

    }
    if (($ExchangeSenderMemberOfException.count -ne 0) -or ($ExchangeSenderException.count -ne 0) ) {
        
        $ConfigObjectResult["Excluded Exchange Groups"] += "$ExchangeSenderMemberOfException $ExchangeSenderException " 

    }
    Return $ConfigObjectResult
                               
}

function Set-SharePointNotAllLocationEnabledConfigObject {
    Param
    (
        $ConfigObjectResult,
        $CompliancePolicy
     
    )
    $SharePointLocation = $CompliancePolicy.SharePointLocation
    $SharePointLocationException = $CompliancePolicy.SharePointLocationException
    $SharePointOnPremisesLocationException = $CompliancePolicy.SharePointOnPremisesLocationException

    if (((@($SharePointLocation) -like 'All').Count -lt 1)) {  
        if (@($SharePointLocation).count -ne 0) {
            
            $ConfigObjectResult["Included SP Sites"] += "$SharePointLocation " 
        }
    }
    
    if (($SharePointLocationException.count -ne 0) -or ($SharePointOnPremisesLocationException.count -ne 0)) { 
        
        $ConfigObjectResult["Excluded SP Sites"] += "$SharePointLocationException $SharePointOnPremisesLocationException " 
    }
    
    Return $ConfigObjectResult
                               
}

function Set-TeamsNotAllLocationEnabledConfigObject { 
    Param
    (
        $ConfigObjectResult,
        $CompliancePolicy
    )

    $TeamsLocation = $CompliancePolicy.TeamsLocation
    $TeamsLocationException = $CompliancePolicy.TeamsLocationException
   
    if (((@($TeamsLocation) -like 'All').Count -lt 1)) {  
        if (@($TeamsLocation).count -ne 0) {
            
            $ConfigObjectResult["Included Teams Account"] += "$TeamsLocation" 
        }
    }
   
    if (($TeamsLocationException.count -ne 0)) {
        
        $ConfigObjectResult["Excluded Teams Account"] += "$TeamsLocationException" 
    }
    Return $ConfigObjectResult
                               
}
function Set-OneDriveNotAllLocationEnabledConfigObject {
    Param
    (
        $ConfigObject,
        $PolicySensitiveType,
        $CompliancePolicy
        
    )
    $OneDriveLocation = $CompliancePolicy.OneDriveLocation
    $OneDriveLocationException = $CompliancePolicy.OneDriveLocationException
    $ExceptIfOneDriveSharedByMemberOf = $CompliancePolicy.ExceptIfOneDriveSharedByMemberOf

    if (((@($OneDriveLocation) -like 'All').Count -lt 1)) {  
        if (@($OneDriveLocation).count -ne 0) {
            
            $ConfigObjectResult["Included OneDrive Account"] += "$OneDriveLocation" 
        }

    }
 
    if (($OneDriveLocationException.count -ne 0) -or ($ExceptIfOneDriveSharedByMemberOf.count -ne 0)) {
        
        $ConfigObjectResult["Excluded OneDrive Account"] += "$OneDriveLocationException $ExceptIfOneDriveSharedByMemberOf" 
        
    }
    Return $ConfigObjectResult
                               
}
function Set-DevicesNotAllLocationEnabledConfigObject {
    Param
    (
        $ConfigObject,
        $PolicySensitiveType,
        $CompliancePolicy
        
    )
    $EndpointDlpLocation = $CompliancePolicy.EndpointDlpLocation
    $EndpointDlpLocationException = $CompliancePolicy.EndpointDlpLocationException

    if (((@($EndpointDlpLocation) -like 'All').Count -lt 1)) {  
        if (@($EndpointDlpLocation).count -ne 0) {
            
            $ConfigObjectResult["Included Devices User/Groups"] += "$EndpointDlpLocation" 
        }

    }
 
    if (($EndpointDlpLocationException.count -ne 0)) {
        
        $ConfigObjectResult["Excluded Devices User/Groups"] += "$EndpointDlpLocationException" 
        
    }
    Return $ConfigObjectResult
                               
}