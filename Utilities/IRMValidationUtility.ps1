
using module "..\MCCA.psm1"
<#
 This function returns list of parent labels and sublabels
#>

Function Get-IRMConfigurationPolicy {
    Param(
        $Config,
        $Templates,
        $LogFile
    )
    $ConfigObjectList = @()
    try {
        $AnyPolicyEnabled = $false
        $IRMPolicy = @()
        foreach($Template in $templates)
        {
            $IRMPolicy += $Config["GetInsiderRiskPolicy"] | Where-Object { $_.InsiderRiskScenario -eq $Template }

        }

        foreach ($Policy in $IRMPolicy) {
            if ($($Policy.Mode) -eq "Enable") {
                if ($AnyPolicyEnabled -eq $false) {
                    $AnyPolicyEnabled = $true
                }
                
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "Policy"
                $ConfigObject.ConfigItem = "$($Policy.Name)"

                $UsergroupsEnabled = ""
                $ExchangeLocation = $Policy.ExchangeLocation
                foreach ($Location in $ExchangeLocation) {
                    if ($UsergroupsEnabled -eq "") {
                        $UsergroupsEnabled += "$Location"
                    }
                    else {
                        $UsergroupsEnabled += ", $Location"
                    }
                }
                if ($($Policy.InsiderRiskScenario) -eq "HighValueEmployeeDataLeak") {
                    $PolicyGroups = $Policy.CustomTags
                    foreach ($PolicyGroup in $PolicyGroups) {
                        $Group = $PolicyGroup.Split("""")#The policy group details come as string hence parsing to get group name
                        if ($UsergroupsEnabled -eq "") {
                            $UsergroupsEnabled += "$($Group[3])"
                        }
                        else {
                            $UsergroupsEnabled += ", $($Group[3])"
                        }
                    }
                }
                $ConfigObject.ConfigData = "$UsergroupsEnabled"

                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")            
                $ConfigObjectList += $ConfigObject
            }
        }

        if ($AnyPolicyEnabled -eq $false) {
            $ConfigObject = [MCCACheckConfig]::new()
            $ConfigObject.Object = "Policy"
            $ConfigObject.ConfigItem = "<B>No active policy defined<B>"
            $ConfigObject.ConfigData = ""
            $ConfigObject.SetResult([MCCAConfigLevel]::OK, "Fail")            
            $ConfigObjectList += $ConfigObject
        }
        
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue while running MCCA. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    return $ConfigObjectList
}

