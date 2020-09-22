
using module "..\MCCA.psm1"
<#
 This function returns list of parent labels and sublabels
#>

Function Get-RoleGroupwithMembers {
    Param(
        $RoleGroups,
        $LogFile,
        $Role
    )
    
    $ConfigObjectList = @()
    try {
        $RoleGroupName = ""
        if ( $null -eq $RoleGroups) {
            $ConfigObject = [MCCACheckConfig]::new()
            $ConfigObject.ConfigItem = "$Role"
            $ConfigObject.ConfigData = "No Role Group with any Members"
            $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
            $ConfigObjectList += $ConfigObject
        }
        else {
            foreach ($RoleGroup in $RoleGroups ) {
                if ($RoleGroupName -ne "") {
                    $RoleGroupName += ", $($RoleGroup.Name)"
                }
                else {
                    $RoleGroupName = $($RoleGroup.Name)
                }
            }
            $ConfigObject = [MCCACheckConfig]::new()
            $ConfigObject.ConfigItem = "$Role"
            $ConfigObject.ConfigData = $($RoleGroupName)
            $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass")  
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

