#Requires -Version 5.1

<#
	.SYNOPSIS
		MCCA - Microsoft Compliance Configuration Analyzer (MCCA)

	.DESCRIPTION

	.NOTES
		Neha Pandey
		Senior Software Engineer - Microsoft
              
        Kritika Mishra
        Software Engineer - Microsoft
            


        Output report uses open source components for HTML formatting
        - bootstrap - MIT License - https://getbootstrap.com/docs/4.0/about/license/
        - fontawesome - CC BY 4.0 License - https://fontawesome.com/license/free
        
        ############################################################################

        This sample script is not supported under any Microsoft standard support program or service. 
        This sample script is provided AS IS without warranty of any kind. 
        Microsoft further disclaims all implied warranties including, without limitation, any implied 
        warranties of merchantability or of fitness for a particular purpose. The entire risk arising 
        out of the use or performance of the sample script and documentation remains with you. In no
        event shall Microsoft, its authors, or anyone else involved in the creation, production, or 
        delivery of the scripts be liable for any damages whatsoever (including, without limitation, 
        damages for loss of business profits, business interruption, loss of business information, 
        or other pecuniary loss) arising out of the use of or inability to use the sample script or
        documentation, even if Microsoft has been advised of the possibility of such damages.

        ############################################################################    

	.LINK
        about_functions_advanced

#>
[bool] $global:ErrorOccurred = $false

function Get-MCCADirectory {
    <#

        Gets or creates the MCCA directory in AppData
        
    #>
    If ($IsWindows) {
        $Directory = "$($env:LOCALAPPDATA)\Microsoft\MCCA"
    }
    elseif ($IsLinux -or $IsMac) {
        $Directory = "$($env:HOME)/MCCA"
    }
    else {
        $Directory = "$($env:LOCALAPPDATA)\Microsoft\MCCA"
    }
	
    If (Test-Path $Directory) {
        Return $Directory
    } 
    else {
        mkdir $Directory | out-null
        Return $Directory
    }
}

Function Invoke-MCCAConnections {
    Param
    (
        [String]$LogFile
    )
   
    
    try {

        try
        {
            $ExchangeVersion = (Get-InstalledModule -name "ExchangeOnlineManagement" -ErrorAction:SilentlyContinue | Sort-Object Version -Desc)[0].Version
        }
        catch
        {
            # EOM(Exchange Online Management) is not installed
            $ExchangeVersion = "Error"
            write-host "$(Get-Date) Exchange Online Management module is not installed. Installing.."
            Install-Module -Name "ExchangeOnlineManagement" -force
        }
    
        if($ExchangeVersion -eq "Error")
        {
            $ExchangeVersion = (Get-InstalledModule -name "ExchangeOnlineManagement" | Sort-Object Version -Desc)[0].Version
        }
        
        if("$ExchangeVersion" -ne "2.0.3")
        {
            write-host "$(Get-Date) Your Exchange Online Management module is not updated. Updating.."
            Update-Module -Name "ExchangeOnlineManagement" -RequiredVersion 2.0.3
        }

        $userName = Read-Host -Prompt 'Input the user name' -ErrorAction:SilentlyContinue
        $InfoMessage = "Connecting to Exchange Online (Modern Module).."
        Write-Host "$(Get-Date) $InfoMessage"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Connect-ExchangeOnline -Prefix EXOP -UserPrincipalName $userName -ShowBanner:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue in connecting to Exchange Online. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }

    try {
        $InfoMessage = "Connecting to Security & Compliance Center"
        Write-Host "$(Get-Date) $InfoMessage"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Connect-IPPSSession -UserPrincipalName $userName -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue in connecting to Security & Compliance Center. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
        throw 'There was an issue in connecting to Security & Compliance Center. Please try running the tool again after some time.'
    }
}

enum CheckType {
    ObjectPropertyValue
    PropertyValue
}

[Flags()]
enum MCCAService {
    DLP = 1
    OATP = 2
}

enum MCCAConfigLevel {
    None = 0
    Recommendation = 4
    Ok = 5
    Informational = 10
    TooStrict = 15
}

enum MCCAResult {
    Pass = 1
    Recommendation = 2
    Fail = 3
}

Class MCCACheckConfig {

    MCCACheckConfig() {
        # Constructor

        $this.Results += New-Object -TypeName MCCACheckConfigResult -Property @{
            Level = [MCCAConfigLevel]::Recommendation
        }

        $this.Results += New-Object -TypeName MCCACheckConfigResult -Property @{
            Level = [MCCAConfigLevel]::Ok
        }

        $this.Results += New-Object -TypeName MCCACheckConfigResult -Property @{
            Level = [MCCAConfigLevel]::Informational 
        }

        $this.Results += New-Object -TypeName MCCACheckConfigResult -Property @{
            Level = [MCCAConfigLevel]::TooStrict
        }

    }

    # Set the result for this mode
    SetResult([MCCAConfigLevel]$Level, $Result) {
        ($this.Results | Where-Object { $_.Level -eq $Level }).Value = $Result
        # The level of this configuration should be its strongest result (e.g if its currently Ok and we have a Informational pass, we should make the level Informational)
        if ($Result -eq "Pass" -and ($this.Level -lt $Level -or $this.Level -eq [MCCAConfigLevel]::None)) {
            $this.Level = $Level
        } 
        elseif ($Result -eq "Fail" -and ($Level -eq [MCCAConfigLevel]::Recommendation -and $this.Level -eq [MCCAConfigLevel]::None)) {
            $this.Level = $Level
        }

    }

    $Check
    $Object
    $ConfigItem
    $ConfigData
    $InfoText
    [string]$RemediationAction = ""
    [array]$Results
    [MCCAConfigLevel]$Level
}

Class MCCACheckConfigResult {
    [MCCAConfigLevel]$Level = [MCCAConfigLevel]::Ok
    $Value
}
Class MCCARemediationInfo {
    [bool]$RemediationAvailable = $false
    [string]$RemediationText = ""
}

Class MCCACheck {
    <#

        Check definition

        The checks defined below allow contextual information to be added in to the report HTML document.
        - Control               : A unique identifier that can be used to index the results back to the check
        - Area                  : The area that this check should appear within the report
        - PassText              : The text that should appear in the report when this 'control' passes
        - FailRecommendation    : The text that appears as a title when the 'control' fails. Short, descriptive. E.g "Do this"
        - Importance            : Why this is important
        - ExpandResults         : If we should create a table in the callout which points out which items fail and where
        - ObjectType            : When ExpandResults is set to, For Object, Property Value checks - what is the name of the Object, e.g a Spam Policy
        - ItemName              : When ExpandResults is set to, what does the check return as ConfigItem, for instance, is it a Transport Rule?
        - DataType              : When ExpandResults is set to, what type of data is returned in ConfigData, for instance, is it a Domain?    

    #>

    [Array] $Config = @()
    [string] $Control
    [string] $ParentArea
    [String] $Area
    [String] $Name
    [String] $PassText
    [String] $FailRecommendation
    [Boolean] $ExpandResults = $false
    [String] $ObjectType
    [String] $ItemName
    [String] $DataType
    [String] $Importance
    [MCCAService]$Services = [MCCAService]::DLP
    [CheckType] $CheckType = [CheckType]::PropertyValue
    [MCCARemediationInfo] $MCCARemediationInfo
    [string] $LogFile 
    $Links
    $MCCAParams

    [MCCAResult] $Result = [MCCAResult]::Pass
    [int] $FailCount = 0
    [int] $PassCount = 0
    [int] $InfoCount = 0
    [Boolean] $Completed = $false
    
    # Overridden by check
    GetResults($Config) { }

    AddConfig([MCCACheckConfig]$Config) {
        $this.Config += $Config

        $this.FailCount = @($this.Config | Where-Object { $_.Level -eq [MCCAConfigLevel]::None }).Count
        $this.PassCount = @($this.Config | Where-Object { $_.Level -eq [MCCAConfigLevel]::Ok -or $_.Level -eq [MCCAConfigLevel]::Informational }).Count
        $this.InfoCount = @($this.Config | Where-Object { $_.Level -eq [MCCAConfigLevel]::Recommendation }).Count

        If ($this.FailCount -eq 0 -and $this.InfoCount -eq 0) {
            $this.Result = [MCCAResult]::Pass
        }
        elseif ($this.FailCount -eq 0 -and $this.InfoCount -gt 0) {
            $this.Result = [MCCAResult]::Recommendation
        }
        else {
            $this.Result = [MCCAResult]::Fail    
        }
        
       

    }

    # Run
    Run($Config) {
        Write-Host "$(Get-Date) Analysis - $($this.Area) - $($this.Name)"
        
        $this.GetResults($Config)

        # If there is no results to expand, turn off ExpandResults
        if ($this.Config.Count -eq 0) {
            $this.ExpandResults = $false
        }

        
    }

}

Class MCCAOutput {

    [String]    $Name
    [Boolean]   $Completed = $False
    $VersionCheck
    $DefaultOutputDirectory
    $Result

    # Function overridden
    RunOutput($Checks, $Collection) {

    }

    Run($Checks, $Collection) {

        $this.RunOutput($Checks, $Collection)

        $this.Completed = $True
    }

}
Class RemediationAction {

    [String]    $Name
    [Boolean]   $Completed = $False
    $VersionCheck
    $DefaultOutputDirectory
    $Result

    # Function overridden
    RunOutput($Checks, $Collection) {

    }

    Run($Checks, $Collection) {

        $this.RunOutput($Checks, $Collection)

        $this.Completed = $True
    }

}

Function Get-MCCACheckDefs {
    Param
    (
        [string]$LogFile,
        $MCCAParams,
        $Collection
    )

    $Checks = @()

    # Load individual check definitions
    $CheckFiles = Get-ChildItem "$PSScriptRoot\Checks"

    # DLP check file full name
    $DLPCheckFileName = $null
    
    #Setting DLP check file name
    ForEach ($CheckFile in $CheckFiles) {
        if (($CheckFile.BaseName -match '^check-(.*)$') -and ($matches[1] -like "DLP")) {
            $DLPCheckFileName = $CheckFile.FullName
        }
    }
    
    
    #Creating DLP check objects for each improvement actions

    #read xml doc
    if($($Collection["GetRequiredSolution"]) -icontains "DLP"){
    [xml]$CheckData = Get-Content "$PSScriptRoot\DLPImprovementActions\ActionsInformation.xml"
    if ($null -eq $CheckData -or $CheckData -eq "") {
        Write-Host "$(Get-Date) ActionsInformation.xml file does not exist/is corrupt in $PSScriptRoot\DLPImprovementActions\ActionsInformation.xml." -ForegroundColor Orange           
    }

    if ($null -ne $DLPCheckFileName -or $DLPCheckFileName -ne "") {
        Write-Verbose "Importing DLP"
        . $DLPCheckFileName
        foreach ($Item in $CheckData.ImprovementActions.ActionItem) {
            #List of SIT
            $ListOfSIT = @()
            $AllSITS = $Item.SITs.SIT

            
            #Adding custom SITS
      <#      if($($Collection["GetDLPCustomSIT"]) -ne "Error")
            {
                $CustomSIT = $($Collection["GetDLPCustomSIT"]).Name
                foreach ($sit in $CustomSIT) {
                    $ListOfSIT += $sit
                }
            }
    #>

            if($($Collection["GetOrganisationRegion"]) -eq "Error")
            {
                foreach ($sit in $AllSITS) {
                    $ListOfSIT += $sit.InnerText
                }
            }
            else
            {
                foreach ($sit in $AllSITS) {
                    if($($Collection["GetOrganisationRegion"]) -contains $($sit.Geo))
                    {
                        $ListOfSIT += $sit.InnerText
                    }
                }

            }
    
            #Hash table of links
            $LinksInfo = @{}
            $AllLinks = $Item.Links.Link
            foreach ($url in $AllLinks) {
                $LinksInfo[$url.LinkText] = $url.ActualURL
            }
            $InfoParams = @{}
            $InfoParams["Control"] = $Item.CheckName
            $InfoParams["ParentArea"] = $Item.ParentArea
            $InfoParams["Area"] = $Item.Area
            $InfoParams["Name"] = $Item.Name
            $InfoParams["RemediationPolicyName"] = $Item.RemediationPolicyName
            $InfoParams["PassText"] = $Item.PassText
            $InfoParams["FailRecommendation"] = $Item.FailRecommendation
            $InfoParams["Importance"] = $Item.Importance
            $InfoParams["SIT"] = $ListOfSIT
            $InfoParams["Links"] = $LinksInfo
            $Check = New-Object -TypeName "DLP" -ArgumentList $InfoParams
            # Set the MCCAParams
            $Check.MCCAParams = $MCCAParams
            $Check.LogFile = $LogFile
    
            $Checks += $Check
        }
    
    }
}


    # Creating Non-DLP check objects for each improvement actions
    ForEach ($CheckFile in $CheckFiles) {
        if ($CheckFile.BaseName -match '^check-(.*)$' -and ($matches[1] -notlike "DLP")) {
            $solutioname=$matches[1]
            $length=$solutioname.length
            $solutioname=$solutioname.substring(0,$length-3)
            
            if (($null -ne $($Collection["GetRequiredSolution"]))-and ($($Collection["GetRequiredSolution"]) -icontains "$solutioname")) {
            Write-Verbose "Importing $($matches[1])"
            . $CheckFile.FullName
            $Check = New-Object -TypeName $matches[1]
            # Set the MCCAParams
            $Check.MCCAParams = $MCCAParams
            $Check.LogFile = $LogFile
            $Checks += $Check
        }
        
    }
    }
    $Checks = $Checks | Sort-Object -Property @{ expression='ParentArea' ; descending=$true}, @{expression='Area' ;descending=$false}

    Return $Checks
}

Function Get-MCCARemediationAction {
    Param
    (
        $VersionCheck
    )

    $RemediationActions = @()
    # Load individual check definitions
    $RemediationActionOutputFiles = Get-ChildItem "$PSScriptRoot\Remediation"

    ForEach ($RemediationActionOutputFile in $RemediationActionOutputFiles) {
        if ($RemediationActionOutputFile.BaseName -match '^remediation(.*)$') {
            Write-Verbose "Importing $($matches[1])"
            . $RemediationActionOutputFile.FullName
            $RemediationAction = New-Object -TypeName $matches[1]

            # For default output directory
            $RemediationAction.DefaultOutputDirectory = Get-MCCADirectory

            # Provide versioncheck
            $RemediationAction.VersionCheck = $VersionCheck

            $RemediationActions += $RemediationAction
        }

    }

    Return $RemediationActions
}
Function Get-MCCAOutputs {
    Param
    (
        $VersionCheck,
        $Modules,
        $Options
    )

    $Outputs = @()

    # Load individual check definitions
    $OutputFiles = Get-ChildItem "$PSScriptRoot\Outputs"

    ForEach ($OutputFile in $OutputFiles) {
        if ($OutputFile.BaseName -match '^output-(.*)$') {
            # Determine if this type should be loaded
            If ($Modules -contains $matches[1]) {
                Write-Verbose "Importing $($matches[1])"
                . $OutputFile.FullName
                $Output = New-Object -TypeName $matches[1]

                # Load any of the options in to the module
                If ($Options) {

                    If ($Options[$matches[1]].Keys) {
                        ForEach ($Opt in $Options[$matches[1]].Keys) {
                            # Ensure this property exists before we try set it and get a null ref error
                            $ModProperties = $($Output | Get-Member | Where-Object { $_.MemberType -eq "Property" }).Name
    
                            If ($ModProperties -contains $Opt) {
                                $Output.$Opt = $Options[$matches[1]][$Opt]
                            }
                            else {
                                Throw("There is no option $($Opt) on output module $($matches[1])")
                            }
                        }
                    }
                }

                # For default output directory
                $Output.DefaultOutputDirectory = Get-MCCADirectory

                # Provide versioncheck
                $Output.VersionCheck = $VersionCheck
                
                $Outputs += $Output
            }

        }
    }

    Return $Outputs
}

# Get DLP settings
Function Get-DataLossPreventionSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetDlpComplianceRule"] = Get-DlpComplianceRule -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $Collection["GetDLPCustomSIT"] = Get-DlpSensitiveInformationType -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage | Where-Object { $_.Publisher -ne "Microsoft Corporation" } 
        $Collection["GetDlpCompliancePolicy"] = Get-DlpCompliancePolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {        
        $Collection["GetDlpComplianceRule"] = "Error"
        $Collection["GetDLPCustomSIT"] = "Error"
        $Collection["GetDlpCompliancePolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Data Loss Prevention information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }

    Return $Collection
}

# Get Information Protection settings
Function Get-InformationProtectionSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetLabel"] = Get-Label -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        try {
            $Collection["GetLabelPolicy"] = Get-LabelPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        }
        catch {
            $Collection["GetLabelPolicy"] = "Error"
        }
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetLabel"] = "Error"
        $Collection["GetLabelPolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Information Protection information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
         
    }
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetAutoSensitivityLabelPolicy"] = Get-AutoSensitivityLabelPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetAutoSensitivityLabelPolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching AutoSensitivity Label Policy information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetIRMConfiguration"] = Get-EXOPIRMConfiguration -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetIRMConfiguration"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching IRM Configuration information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    
    }
    Return $Collection
}

# Get Communication Compliance settings
Function Get-CommunicationComplianceSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetSupervisoryReviewPolicyV2"] = Get-SupervisoryReviewPolicyV2 -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        try {
            $Collection["GetSupervisoryReviewOverallProgressReport"] = Get-SupervisoryReviewOverallProgressReport -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        
        }
        catch {
            $Collection["GetSupervisoryReviewOverallProgressReport"] = "Error"
        }
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetSupervisoryReviewPolicyV2"]  = "Error"
        $Collection["GetSupervisoryReviewOverallProgressReport"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Communication Compliance information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
         
    }
    Return $Collection
}

# Get Information Governance settings
Function Get-InformationGovernanceSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetRetentionCompliancePolicy"] = Get-RetentionCompliancePolicy -DistributionDetail -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetRetentionComplianceRule"] = Get-RetentionComplianceRule -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetComplianceTag"] = Get-ComplianceTag -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetRetentionCompliancePolicy"] = "Error"
        $Collection["GetRetentionComplianceRule"] = "Error"
        $Collection["GetComplianceTag"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Retention Compliance information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
         
    }
    Return $Collection
}

# Get Audit settings
Function Get-AuditSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetAdminAuditLogConfig"] = Get-EXOPAdminAuditLogConfig -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetAdminAuditLogConfig"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Audit Configuration information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
    }
    Return $Collection
}
#get eDiscovery
Function Get-eDiscoverySettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetComplianceCase"] = Get-ComplianceCase -CaseType AdvancedEdiscovery -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetComplianceCaseCore"] = Get-ComplianceCase  -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetComplianceCase"] = "Error"
        $Collection["GetComplianceCaseCore"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Audit Configuration information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
    }
    Return $Collection
}

#Get Insider Risk Management Settings
Function Get-InsiderRiskManagementSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetInsiderRiskPolicy"] = Get-InsiderRiskPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetInsiderRiskPolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Insider Risk Management information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
    }
    Return $Collection
}

# Get Accepted Domains
Function Get-AcceptedDomains {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["AcceptedDomains"] = Get-EXOPAcceptedDomain -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue in fetching tenant name information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
    }
    Return $Collection
}

#Get Alert Policies
Function Get-AlertPolicies {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetProtectionAlert"] = Get-ProtectionAlert | Where-Object { $_.Severity -eq "High" } -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetProtectionAlert"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Alert Policies Configuration information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
    }
    Return $Collection
}


#Get Organisation Region
Function Get-OrganisationRegion {
    Param(
        $Collection,
        [string]$LogFile,
        [System.Collections.ArrayList] $GeoList
    )
    
    
        try {
            [System.Collections.ArrayList]$WarnMessage = @()
            [System.Collections.ArrayList] $RegionNamesList = @()
            $Collection["GetOrganisationConfig"] = Get-EXOPOrganizationConfig -ErrorAction:SilentlyContinue
            
            if($($GeoList.Count) -gt 0)
            {
               $Collection["GetOrganisationRegion"] = $GeoList
               $Collection["GetOrganisationRegion"].add("INTL") | out-null
            }
            else
            {
                $RegionsList = $Collection["GetOrganisationConfig"].AllowedMailboxRegions 
                foreach ($region in $RegionsList) {
                $RegionName = $($region.Split("="))[0]
                $RegionName = $RegionName.ToUpper()
                $RegionNamesList.add($RegionName) | Out-Null
                }
                $Collection["GetOrganisationRegion"] = $RegionNamesList
                $Collection["GetOrganisationRegion"].add("INTL") | out-null
            }
            Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        
            
        }
        catch {
            $Collection["GetOrganisationConfig"] = "Error"
            if($($GeoList.Count) -gt 0)
            {
               $Collection["GetOrganisationRegion"] = $GeoList
               $Collection["GetOrganisationRegion"].add("INTL") | out-null
            }
            else
            {
                $Collection["GetOrganisationRegion"] = "Error"
                Write-Host "Warning:$(Get-Date) There was an issue in fetching your tenant's geolocation. The generated report will have recommendations for all geos across the globe." -ForegroundColor:Yellow
            
            }
            $ErrorMessage = $_.ToString()
            $StackTraceInfo = $_.ScriptStackTrace
            Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue      
        }
    
        
    Return $Collection
}

#Get Solution Config
Function Get-PersonalizedSolution {
    Param(
        $Collection,
        [string]$LogFile,
        [System.Collections.ArrayList] $SolutionList
    )
       
            [System.Collections.ArrayList] $SolutionsList = @()
            if($($SolutionList.Count) -gt 0)
            {
               $Collection["GetRequiredSolution"] = $SolutionList
               $Collection["GetRequiredSolution"].add("INTL") | out-null
            }
            else
            {
                $SolutionTable = Get-SolutionTable
                [int] $count = 1
                while ($count -le 8) {
                 $SolutionList.add($($($SolutionTable[$count]).Code)) |out-null
                   $count = $count + 1
                  }

                $Collection["GetRequiredSolution"] = $SolutionsList
                $Collection["GetRequiredSolution"].add("INTL") | out-null
            }
      Return $Collection
}
               
# Get user configurations
Function Get-MCCACollection {
    Param
    (
        [String]$LogFile,
        [System.Collections.ArrayList] $GeoList,
        [System.Collections.ArrayList] $SolutionList
    )
    $Collection = @{}

    [MCCAService]$Collection["Services"] = [MCCAService]::DLP
    try{
        Write-EXOPAdminAuditLog -Comment "MCCA Started at- $(Get-Date)"

    }catch{
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    if($SolutionList -icontains "DLP")
    {$InfoMessage = "Getting DLP Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-DataLossPreventionSettings -Collection $Collection -LogFile $LogFile}

    if($SolutionList -icontains "IP")
    {$InfoMessage = "Getting Information Protection Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-InformationProtectionSettings -Collection $Collection -LogFile $LogFile}

    if($SolutionList -icontains "CC")
    {$InfoMessage = "Getting Communication Compliance Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-CommunicationComplianceSettings -Collection $Collection -LogFile $LogFile}
    
    if(($SolutionList -icontains "IG") -or ($SolutionList -icontains "RM"))
    {$InfoMessage = "Getting Information Governance Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-InformationGovernanceSettings -Collection $Collection -LogFile $LogFile}
    
    if($SolutionList -icontains "Audit" )
    {$InfoMessage = "Getting Audit Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-AuditSettings -Collection $Collection -LogFile $LogFile}

    if($SolutionList -icontains "eDiscovery")
    {$InfoMessage = "Getting eDiscovery Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-eDiscoverySettings -Collection $Collection -LogFile $LogFile}

    if($SolutionList -icontains "IRM")
    {$InfoMessage = "Getting Insider Risk Management Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-InsiderRiskManagementSettings -Collection $Collection -LogFile $LogFile}

    $InfoMessage = "Getting Accepted Domains"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-AcceptedDomains -Collection $Collection -LogFile $LogFile
    
    $InfoMessage = "Getting Alert Policies Settings"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-AlertPolicies -Collection $Collection -LogFile $LogFile

    $InfoMessage = "Getting Organization's region information"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-OrganisationRegion -GeoList $GeoList -Collection $Collection -LogFile $LogFile

    $InfoMessage = "Getting Organization's solution preference information"
    Write-Host "$(Get-Date) $InfoMessage"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Collection = Get-PersonalizedSolution -SolutionList $SolutionList -Collection $Collection -LogFile $LogFile

    Return $Collection
}


Function Get-MCCAReport {
<#
    
        .SYNOPSIS
            The Microsoft Compliance Configuration Analyzer (MCCA)

        .DESCRIPTION
            Microsoft Compliance Configuration Analyzer (MCCA)

            The Get-MCCAReport command generates a HTML report highlighting known issues in your compliance configurations in achieving data protection guidelines and recommends best practices to follow.

            Output report uses open source components for HTML formatting:
            - Bootstrap - MIT License https://getbootstrap.com/docs/4.0/about/license/
            - Fontawesome - CC BY 4.0 License - https://fontawesome.com/license/free

       
        .PARAMETER NoVersionCheck
            Prevents MCCA from determining if it's running the latest version. It's always very important to be running the latest
            version of MCCA. We will change guidelines as the product and the recommended practices article changes. Not running the
            latest version might provide recommendations that are no longer valid.

		.PARAMETER Geo 
            This will generate a report based on the geolocations entered by you.You need to input appropriate numbers from the following list corresponding to the regions. 
			Input	Region
				1	Asia-Pacific
				2	Australia
				3	Canada
				4	Europe (excl. France) / Middle East / Africa
				5	France
				6	India
				7	Japan
				8	Korea
				9	North America (excl. Canada)
				10	South America
				11	South Africa
				12	Switzerland
				13	United Arab Emirates
				14	United Kingdom


		.PARAMETER Solution 
            This will generate a report only for the solutions entered by you. You need to input appropriate numbers from the following list corresponding to the solution. 
			Input	Solution
				1	Data Loss Prevention
				2	Information Protection
				3	Information Governance
				4	Records Management
				5	Communication Compliance
				6	Insider Risk Management
				7	Audit
				8	eDiscovery

        .PARAMETER Collection
            Internal only.
        .EXAMPLE
            Get-MCCAReport
			This will generate a customized report based on the geolocation of your tenant. If an error occurs while fetching your tenant's geolocation, you will get a report covering all supported geolocations.
		.EXAMPLE
            Get-MCCAReport -Geo @(1,7)
			This will generate a customized report based on the geolocations entered by you. 
		.EXAMPLE
			Get-MCCAReport -Solution @(1,7)
			This will generate a customized report for the solutions entered by you. 
		.EXAMPLE
		    Get-MCCAReport -Solution @(1,7) -Geo @(9)
			This will generate a report only on for the solutions entered by you and based on the regions you have selected. 

    
#>
    Param(
        [CmdletBinding()]
        [Switch]$NoVersionCheck,    
        [System.Collections.ArrayList] $Geo=@(),
        [System.Collections.ArrayList] $Solution=@(),
        $Collection
    )
    
    $OutputDirectoryName = Get-MCCADirectory
    $LogDirectory = "$OutputDirectoryName\Logs"
    $FileName = "MCCA-$(Get-Date -Format 'yyyyMMddHHmmss').log"
    $LogFile = "$LogDirectory\$FileName"
    #Creating the logfiles folder if not present
    if ($(Test-Path -Path $LogDirectory) -eq $false) {
        New-Item -Path $LogDirectory -ItemType Directory -ErrorAction:SilentlyContinue | Out-Null
        #Creating the logfile
        New-Item -Path $LogFile -ItemType File -ErrorAction:SilentlyContinue | Out-Null
    }
    else {
        New-Item -Path $LogFile -ItemType File -ErrorAction:SilentlyContinue | Out-Null
    }
    #Check if log file exists
    if ($(Test-Path -Path $LogFile) -eq $False) {
        Write-Host "$(Get-Date) Log file cannot be created." -ForegroundColor:Red
    }
    Write-Log -MachineInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    if(($(Get-GeoAcceptance -Geo $Geo) -eq $false ) -and ($(Get-SolutionAcceptance -Solution $Solution) -eq $false))
    {     
        Show-GeoOptions
        Show-SolutionOptions 
        return
    }
    #Get actual region names
    
    [System.Collections.ArrayList] $GeoList = @()
    if(($(Get-GeoAcceptance -Geo $Geo) -eq $false ))
    {
               
        Show-GeoOptions 
        return
    }
    else
    {
    #Number To Region Mapping 
    $NumberToRegionMapping = Get-NumberRegionMappingHashTable

    #Mapping numbers to the actual region
    foreach ($RegionNumber in $Geo)
    {
        [string] $RegionName = $NumberToRegionMapping[$RegionNumber].Code
        $GeoList.add($RegionName) | out-null  
    }
    }

        #Get actual region names

        [System.Collections.ArrayList] $SolutionList = @()
        if($(Get-SolutionAcceptance -Solution $Solution) -eq $false)
        {
                  
            Show-SolutionOptions 
            return
        }
        else
        {
        $ShowSolutionList =""
        $SolutionTable = Get-SolutionTable
        if($Solution.count -gt 0)
        {foreach ($count in $Solution)
        {
            [string] $Name =  "$($($SolutionTable[$count]).Code)"
            #write-host "$Name"
            $SolutionList.add($Name) | out-null
            $ShowSolutionList += "$($($SolutionTable[$count]).FullName), "
        }
        $ShowSolutionList=$ShowSolutionList.TrimEnd(", ")
        }else {
            
            [int] $count = 1
               while ($count -le 8) {
                $SolutionList.add($($($SolutionTable[$count]).Code))|out-null
                  $count = $count + 1
                 }
                 $ShowSolutionList += "All Solutions"   
        }
    }
    # Easy to use for quick MCCA report to HTML
    If ($NoVersionCheck) {
        $PerformVersionCheck = $False
    }
    Else {
        $PerformVersionCheck = $True
    }
    

    try {
        $Result = Invoke-MCCA -PerformVersionCheck $PerformVersionCheck -Collection $Collection -Output @("HTML") -GeoList $GeoList -SolutionList $SolutionList -LogFile $LogFile -ErrorAction:SilentlyContinue
        $InfoMessage = "Complete! Output is in $($Result.Result)"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host "$(Get-Date) $InfoMessage"
        try{
            Write-EXOPAdminAuditLog -Comment "MCCA Completed at - $(Get-Date)"
    
        }catch{
            $ErrorMessage = $_.ToString()
            $StackTraceInfo = $_.ScriptStackTrace
            Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
        }
        Write-Log -StopInfo -LogFile $LogFile -ErrorAction:SilentlyContinue

    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue in running the tool. Please try running the tool again after some time." -ForegroundColor:Red
        Write-Host "Please refer documentation for more details. If the issue persists, please write to us at MCCAhelp@microsoft.com." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
            
    }
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue
    }
    catch {
        
    }
   
}


Function Invoke-MCCA {
    Param(
        [CmdletBinding()]
        [Boolean]$PerformVersionCheck = $True,     
        $Output,
        $OutputOptions,
        $Collection,
        [System.Collections.ArrayList] $GeoList=@(),
        [System.Collections.ArrayList] $SolutionList=@(),
        [String]$LogFile
    )
    $InfoMessage = "MCCA Started"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue

    # Version check
    If ($PerformVersionCheck) {
        $InfoMessage = "Version Check Started"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        $VersionCheck = Invoke-MCCAVersionCheck 
        $InfoMessage = "Version Check Completed"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    
   
        $InfoMessage = "Establishing Connections"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Invoke-MCCAConnections -LogFile $LogFile
        $InfoMessage = "Connections Established"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    


    # Get the collection in to memory. For testing purposes, we support passing the collection as an object
    If ($Null -eq $Collection) {
        $InfoMessage = "Fetching User Configurations"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        $Collection = Get-MCCACollection -GeoList $GeoList -SolutionList $SolutionList -LogFile $LogFile
        $InfoMessage = "User Configurations Fetched"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }


    # Get the output modules
    $InfoMessage = "Creating Output Objects"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $OutputModules = Get-MCCAOutputs -VersionCheck $VersionCheck -Modules $Output -Options $OutputOptions
    $InfoMessage = "Output Objects Created"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    
    # Get the object of MCCA checks
    $InfoMessage = "Creating Check Objects"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $Checks = Get-MCCACheckDefs -MCCAParams $MCCAParams -Collection $Collection -LogFile $LogFile
    $InfoMessage = "Check Objects Created"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue

    
    
    # Perform checks inside classes/modules
    ForEach ($Check in ($Checks | Sort-Object Area)) {

        # Run DLP checks by default
        if ($check.Services -band [MCCAService]::DLP) {
            $Check.Run($Collection)
        }

       
    }

    # Get the Remedition Steps
    $InfoMessage = "Creating Remediation Objects"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    $RemediationActionModules = Get-MCCARemediationAction -VersionCheck $VersionCheck  
    $InfoMessage = "Remediation Objects Created"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    ForEach ($a in $RemediationActionModules) {

        $a.Run($Checks, $Collection)
       
        
    }

    $TenantGeoLocations = $Collection["GetOrganisationRegion"] | Where-Object { $_ -ne "INTL" }
            if($TenantGeoLocations -ne "Error")
            {
                $RegionString = ""
                $NumberToRegionMapping = Get-NumberRegionMappingHashTable
                foreach ($Region in $TenantGeoLocations) {
                    foreach ($Numbers in $($NumberToRegionMapping.Keys)) {
                        if($($NumberToRegionMapping[$Numbers].Code) -eq $Region)
                        {
                            if($RegionString -eq "")
                            {
                                $RegionString += "$($NumberToRegionMapping[$Numbers].Description)" 
                            }
                            else
                            {
                                $RegionString += ", $($NumberToRegionMapping[$Numbers].Description)" 
                            }
                        }
                    }

                }}else
                {
                    $RegionString = ""
                    $RegionString +="All Geolocations"
                }
    $InfoMessage = "The following report is generated for following solutions:$ShowSolutionList" 
    Write-Host "$(Get-Date) $InfoMessage" -ForegroundColor Yellow
    $InfoMessage = "The following report is for following geolocations:$RegionString"  
    Write-Host "$(Get-Date) $InfoMessage" -ForegroundColor Yellow
    $OutputResults = @()
    $InfoMessage = "Generating Output"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    Write-Host "$(Get-Date) $InfoMessage" -ForegroundColor Green
    # Perform required outputs
    ForEach ($o in $OutputModules) {

        $o.Run($Checks, $Collection)
        $OutputResults += New-Object -TypeName PSObject -Property @{
            Name      = $o.name
            Completed = $o.completed
            Result    = $o.Result
        }

    }

    Return $OutputResults

}


function Invoke-MCCAVersionCheck {
    Param
    (
        $Terminate
    )

    Write-Host "$(Get-Date) Performing MCCA Version check... "

    # When detected we are running the preview release
    $Preview = $False

    try {
        $MCCAVersion = (Get-InstalledModule MCCA -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | Sort-Object Version -Desc)[0].Version
        
    }
    catch {
        $MCCAVersion = (Get-InstalledModule MCCAPreview | Sort-Object Version -Desc)[0].Version
        
        if ($MCCAVersion) {
            $Preview = $True
        }
    }
    
    if ($Preview -eq $False) {
        $PSGalleryVersion = (Find-Module MCCA -Repository PSGallery -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue).Version
    }
    else {
        $PSGalleryVersion = (Find-Module MCCAPreview -Repository PSGallery -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue).Version
    }
    

    If ($PSGalleryVersion -gt $MCCAVersion) {
        $Updated = $False
        If ($Terminate) {
            Throw "MCCA is out of date. Your version is $MCCAVersion and the published version is $PSGalleryVersion. Run Update-Module MCCA ."
        }
        else {
            Write-Host "$(Get-Date) MCCA is out of date. Your version: $($MCCAVersion) published version is $($PSGalleryVersion)"
        }
    }
    else {
        $Updated = $True
    }

    Return New-Object -TypeName PSObject -Property @{
        Updated        = $Updated
        Version        = $MCCAVersion
        GalleryVersion = $PSGalleryVersion
        Preview        = $Preview
    }
}


#Creating log file and directory

#Writing in log file
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [Switch]$IsError = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$IsWarn = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$IsInfo = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$MachineInfo = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$StopInfo = $false,
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage,
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$WarnMessage,
        [Parameter(Mandatory = $false)]
        [string]$InfoMessage,
        [Parameter(Mandatory = $false)]
        [string]$StackTraceInfo,
        [String]$LogFile
    )   

    if ($MachineInfo) {
        $ComputerInfoObj = Get-ComputerInfo 
        $CompName = $ComputerInfoObj.CsName
        $OSName = $ComputerInfoObj.OsName
        $OSVersion = $ComputerInfoObj.OsVersion
        $PowerShellVersion = $PSVersionTable.PSVersion
        try {
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Logging Started" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Start time: $(Get-Date)" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Computer Name: $CompName" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Operating System Name: $OSName" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Operating System Version: $OSVersion" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "PowerShell Version: $PowerShellVersion" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
         
        }
        catch {
            Write-Host "$(Get-Date) The local machine information cannot be logged." -ForegroundColor:Yellow
        }

    }
    if ($StopInfo) {
        try {
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Logging Ended" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "End time: $(Get-Date)" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            
            if($($global:ErrorOccurred) -eq $true)
            {
                Write-Host "Warning:$(Get-Date) The report generated may have reduced information due to errors in running the tool. These errors may occur due to multiple reasons. Please refer documentation for more details." -ForegroundColor:Yellow
            }
         
        }
        catch {
            Write-Host "$(Get-Date) The finishing time information cannot be logged." -ForegroundColor:Yellow
        }
    }
    #Error
    if ($IsError) {
        if($($global:ErrorOccurred) -eq $false)
        {
            $global:ErrorOccurred = $true
        }
        $Log_content = "$(Get-Date) ERROR: $ErrorMessage"
        try {
            $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "TRACE: $StackTraceInfo" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
        }
        catch {
            Write-Host "$(Get-Date) An error event cannot be logged." -ForegroundColor:Yellow  
        }           
    }
    #Warning
    if ($IsWarn) {
        foreach ($Warnmsg in $WarnMessage) {
            $Log_content = "$(Get-Date) WARN: $Warnmsg"
            try {
                $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            }
            catch {
                Write-Host "$(Get-Date) A warning event cannot be logged." -ForegroundColor:Yellow 
            }
        }
    }
    #General
    if ($IsInfo) {
        $Log_content = "$(Get-Date) INFO: $InfoMessage"
        try {
            $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
        }
        catch {
            Write-Host "$(Get-Date) A general event cannot be logged." -ForegroundColor:Yellow 
        }
        
    }
}

# Get the Number Region Mapping HashTable
function Get-NumberRegionMappingHashTable
{
    #Number To Region Mapping 
    $NumberToRegionMapping = @{}
    $NumberToRegionMapping[1] = New-Object -TypeName PSObject -Property @{
        Code      = "APC"
        Description = "Asia-Pacific"
    }
    $NumberToRegionMapping[2] = New-Object -TypeName PSObject -Property @{
        Code      = "AUS"
        Description = "Australia"
    }
    $NumberToRegionMapping[3] = New-Object -TypeName PSObject -Property @{
        Code      = "CAN"
        Description = "Canada"
    }
    $NumberToRegionMapping[4] = New-Object -TypeName PSObject -Property @{
        Code      = "EUR"
        Description = "Europe (excl. France) / Middle East / Africa"
    }
    $NumberToRegionMapping[5] = New-Object -TypeName PSObject -Property @{
        Code      = "FRA"
        Description = "France"
    }
    $NumberToRegionMapping[6] = New-Object -TypeName PSObject -Property @{
        Code      = "IND"
        Description = "India"
    }
    $NumberToRegionMapping[7] = New-Object -TypeName PSObject -Property @{
        Code      = "JPN"
        Description = "Japan"
    }
    $NumberToRegionMapping[8] = New-Object -TypeName PSObject -Property @{
        Code      = "KOR"
        Description = "Korea"
    }
    $NumberToRegionMapping[9] = New-Object -TypeName PSObject -Property @{
        Code      = "NAM"
        Description = "North America (excl. Canada)"
    }
    $NumberToRegionMapping[10] = New-Object -TypeName PSObject -Property @{
        Code      = "LAM"
        Description = "South America"
    }
    $NumberToRegionMapping[11] = New-Object -TypeName PSObject -Property @{
        Code      = "ZAF"
        Description = "South Africa"
    }
    $NumberToRegionMapping[12] = New-Object -TypeName PSObject -Property @{
        Code      = "CHE"
        Description = "Switzerland"
    }
    $NumberToRegionMapping[13] = New-Object -TypeName PSObject -Property @{
        Code      = "ARE"
        Description = "United Arab Emirates"
    }
    $NumberToRegionMapping[14] = New-Object -TypeName PSObject -Property @{
        Code      = "GBR"
        Description = "United Kingdom"
    }
    
    
    return $NumberToRegionMapping
}


#Check if the geo param is in right format
function Get-GeoAcceptance
{
    param (
        $Geo
    )

    $LegitimateGeo = $Geo | Where-Object { ($_ -ge 1) -and ($_ -le 14) }

    return ($($LegitimateGeo.Count) -eq $($Geo.Count))
}

# Display options for the user to choose
function Show-GeoOptions 
{
    Write-Host "Error:$(Get-Date) Please input appropriate numbers from the following list corresponding to the regions for which you wish to customize the report & run the tool again." -ForegroundColor:Red 
    #Number To Region Mapping 
    $NumberToRegionMapping = Get-NumberRegionMappingHashTable
    Write-Host "*******************************************************************************"
    write-host "For Geo Location"
    Write-Host "*******************************************************************************"
    [int] $count = 1
    while ($count -le 14) {
        Write-Host "$count--->$($($NumberToRegionMapping[$count]).Description)"
        $count = $count + 1
    }

    
    Write-Host "*******************************************************************************"
    Write-Host "Example: Get-MCCAReport -Geo @(1,7) -Solution @(1,7)"
    Write-Host "or"
    Write-Host "Get-MCCAReport -Geo @(1,7)"
    Write-Host ""
    Write-Host ""
    
}

function Get-SolutionTable
{
        #Number To Region Mapping 
        $SolutionTable = @{}
        $SolutionTable[1] = New-Object -TypeName PSObject -Property @{
            Code      = "DLP"
            FullName  =  "Data Loss Prevention"
         
        }
        $SolutionTable[2] = New-Object -TypeName PSObject -Property @{
            Code      = "IP"
            FullName  =  "Information Protection"
        }
        $SolutionTable[3] = New-Object -TypeName PSObject -Property @{
            Code      = "IG"
            FullName  =  "Information Governance"
        }
        $SolutionTable[4] = New-Object -TypeName PSObject -Property @{
            Code      = "RM"
            FullName  =  "Records Management"
        }
        $SolutionTable[5] = New-Object -TypeName PSObject -Property @{
            Code      = "CC"
            FullName  =  "Communication Compliance"
        }
        $SolutionTable[6] = New-Object -TypeName PSObject -Property @{
            Code      = "IRM"
            FullName  =  "Insider Risk Management"
        }
        $SolutionTable[7] = New-Object -TypeName PSObject -Property @{
            Code      = "Audit"
            FullName  =  "Audit"
        }
        $SolutionTable[8] = New-Object -TypeName PSObject -Property @{
            Code      = "eDiscovery"
            FullName  = "eDiscovery"
        }

    return $SolutionTable
}


#Check if the geo param is in right format
function Get-SolutionAcceptance
{
    Param (
        $Solution
    )
    
    $ValidSolution = $Solution | Where-Object { ($_ -ge 1) -and ($_ -le 8) }
    return ($($ValidSolution.Count) -eq $($Solution.Count))
}

function Show-SolutionOptions 
{
  
    Write-Host "Error:$(Get-Date) Please input appropriate numbers from the following list corresponding to solution for which you wish to customize the report & run the tool again." -ForegroundColor:Red 
    $SolutionTable = Get-SolutionTable
    Write-Host "*******************************************************************************"
    write-host "Solution"
    Write-Host "*******************************************************************************"
    [int] $count = 1
    while ($count -le 8) {
        Write-Host "$count--->$($($SolutionTable[$count]).FullName)"
        $count = $count + 1
}

    Write-Host "*******************************************************************************"
    Write-Host "Example: Get-MCCAReport -Geo @(1,7) -Solution @(1,7)"
    Write-Host "or"
    Write-Host "Get-MCCAReport -Solution @(1,7)"
    Write-Host ""
    Write-Host ""
    
}
