
using module "..\MCCA.psm1"
<#
 This function returns list of parent labels and sublabels
#>

Function Get-LableCalssification {
    Param(
        $LogFile
    )
    $SubLabels = @{}
    $ParentLabels = @{}
    $ParentSubLabelAssociation = @{}
    $ParentNameForSubLabelAssociation = @{}
    try {
        foreach ($LabelDefined in $Config["GetLabel"]) {
            $Label = $LabelDefined
            if ($($Label.ParentId)) {
                $SubLabels.add($($Label.Name), $Label)
                if ($ParentSubLabelAssociation.ContainsKey($($Label.ParentId))) {
                    $ParentSubLabelAssociation[$($Label.ParentId)].Add($($Label.Name)) #+= $($Label.Name) 
                }
                else {
                    $ParentSubLabelAssociation.add($($Label.ParentId), [System.Collections.ArrayList]@()) #$($Label.Name))
                    $ParentSubLabelAssociation[$($Label.ParentId)].Add($($Label.Name))
                }
            }
            else {
                $ParentLabels.add($($Label.Name), $Label)
            }
              
        }
    
        # Setting parent name for the parent with sublabels by creating a hash table with key as 
        # parent guid and value as parent name.
        if ($($($ParentSubLabelAssociation.Keys).count) -gt 0) {
            foreach ($ParentGUID in $($ParentSubLabelAssociation.Keys)) {
                foreach ($LabelDefined in $Config["GetLabel"]) {
                    if ($($LabelDefined.Guid) -eq $ParentGUID) {
                        $ParentNameForSubLabelAssociation[$ParentGUID] = $LabelDefined.Name
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error:$(Get-Date) There was an issue while running MCCA. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    

    $LabelClassification = New-Object -TypeName psobject
    $LabelClassification | Add-Member -MemberType NoteProperty -Name sublabels -Value $SubLabels
    $LabelClassification | Add-Member -MemberType NoteProperty -Name parentlabels -Value $ParentLabels    
    $LabelClassification | Add-Member -MemberType NoteProperty -Name parentsublabelassociation -Value $ParentSubLabelAssociation       
    $LabelClassification | Add-Member -MemberType NoteProperty -Name parentnameforsublabelassociation -Value $ParentNameForSubLabelAssociation              
    return $LabelClassification
}

