using module "..\MCCA.psm1"
$ExchangePresent = "Exchange"
$SharePointPresent = "SharePoint"
$OneDrivePresent = "OneDrive"


Function Get-RetentionPolicyValidation {
    param (
        $LogFile,
        $Mode
    )
    $ConfigObjectList = @()
    try {

            $ConfigObjectList=@()
            $LabelPolicy =$false
            $PolicyDisabled = $false
            $AnyPolicyEnabled = $false
            $RetentionComplianceRules = $Config["GetRetentionComplianceRule"]
            $RetentionCompliancePolicies = $Config["GetRetentionCompliancePolicy"]
            $GetComplianceTag = $Config["GetComplianceTag"]
            
            $PartialWorkloadsStatus = @{}
            $PartialWorkloadsStatus[$ExchangePresent] = $false
            $PartialWorkloadsStatus[$SharePointPresent] = $false
            $PartialWorkloadsStatus[$OneDrivePresent] = $false
           
            foreach ( $RetentionCompliancePolicy in $RetentionCompliancePolicies) {
                    
                    $PolicyName = $($RetentionCompliancePolicy.Name)
                    $ConfigObject = [MCCACheckConfig]::new()
                    $ConfigObject.Object = "$PolicyName"
                    $PolicyConfigData = $null

                    if($Mode -eq "Publish")
                    {
                    $RetentionCompliancePolicyRules = $RetentionComplianceRules | Where-Object {($_.Policy -ieq $($RetentionCompliancePolicy.ExchangeObjectId)) -and ($_.Disabled -eq $false) -and($_.PublishComplianceTag -ne "")}
                    foreach( $RetentionCompliancePolicyRule in $RetentionCompliancePolicyRules )
                    {   $PublishComplianceTag = $RetentionCompliancePolicyRule.PublishComplianceTag   
                        $PublishComplianceTag = $($PublishComplianceTag.Split(","))[1] 
                      
                        $GetLabel= $GetComplianceTag | Where-Object{ ($_.Name -eq $PublishComplianceTag) }
                        if( -not (($GetLabel.HasRetentionAction -eq $true) -and ($GetLabel.RetentionDuration -eq "Unlimited")))
                        {if($null -ne $GetLabel)
                        {
                            $LabelPolicy =$true
                            if ($null -eq $PolicyConfigData ) {
                                $PolicyConfigData += "<B>Labels : </B>$($GetLabel.Name)"
                                }
                                else {
                                $PolicyConfigData += ", $($GetLabel.Name)"
                                }
                        }  }  
                    }
                }
                elseif($Mode -eq "Auto")
                {
                    $RetentionCompliancePolicyRules = $RetentionComplianceRules | Where-Object {($_.Policy -ieq $($RetentionCompliancePolicy.ExchangeObjectId)) -and ($_.Disabled -eq $false) -and($_.ApplyComplianceTag -ne "")}
                    foreach( $RetentionCompliancePolicyRule in $RetentionCompliancePolicyRules )
                    {
                        $ApplyComplianceTag = $RetentionCompliancePolicyRule.ApplyComplianceTag   
                         
                      
                        $GetLabel= $GetComplianceTag | Where-Object{ ($_.ExchangeObjectId -eq $ApplyComplianceTag)}
                        if( -not (($GetLabel.HasRetentionAction -eq $true) -and ($GetLabel.RetentionDuration -eq "Unlimited")))
                        {if($null -ne $GetLabel)
                        {

                            $LabelPolicy =$true
                            if ($null -eq $PolicyConfigData ) {
                                $PolicyConfigData += "<B>Labels : </B>$($GetLabel.Name)"
                                }
                                else {
                                $PolicyConfigData += ", $($GetLabel.Name)"
                                }
                        } }   
                    }  
                }
                    $ExchangeLocation = $RetentionCompliancePolicy.ExchangeLocation
                    $SharePointLocation = $RetentionCompliancePolicy.SharePointLocation
                    $OneDriveLocation = $RetentionCompliancePolicy.OneDriveLocation
                    $ModernGroupLocation = $RetentionCompliancePolicy.ModernGroupLocation
                    $PublicFolderLocation = $RetentionCompliancePolicy.PublicFolderLocation
                    $SkypeLocation = $RetentionCompliancePolicy.SkypeLocation
    
                    $WorkloadsStatus= ""
                    if(($RetentionCompliancePolicy.Enabled -eq $true) -and ($null -ne $PolicyConfigData )) 
                    {
                    if(($ExchangeLocation -ne "") )
                    {
                        $WorkloadsStatus+= "Exchange, "
                        $PartialWorkloadsStatus[$ExchangePresent] = $true
                    }
                    if(($SharePointLocation -ne "") )
                    {
                        $WorkloadsStatus += "SharePoint, "
                        $PartialWorkloadsStatus[$SharePointPresent] = $true
                    }
                    if(($OneDriveLocation -ne "") )
                    {
                        $WorkloadsStatus+= "OneDrive, "
                        $PartialWorkloadsStatus[$OneDrivePresent] = $true
                    }
                    if(($ModernGroupLocation -ne "") )
                    {
                        $WorkloadsStatus += "ModernGroup, "
                    }
                    if(($PublicFolderLocation -ne "") )
                    {
                        $WorkloadsStatus += "ExchangePublicFolders, "
                    }
                    if(($SkypeLocation -ne "") )
                    {
                        $WorkloadsStatus += "Skype, "
                    }
                    
                }                
                    $workloadpresent ="<B>Workloads: </B>$WorkloadsStatus"
                    $workloadpresent=$workloadpresent.TrimEnd(", ")
                           
                if (($WorkloadsStatus -ne "")  -and  ($null -ne $PolicyConfigData ) -and ($RetentionCompliancePolicy.Enabled -eq $true)  ) {
                    if ( ($LabelPolicy -eq $true) ) { 
                            $AnyPolicyEnabled =$true
                            $Workload= $workloadpresent
                            $ConfigObject.ConfigData = "$Workload"                                          
                            $ConfigObject.ConfigItem = "$PolicyConfigData"   
                            $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Pass") 
                            $ConfigObjectList += $ConfigObject                              
                        }
                    }
                    elseif(($null -ne $PolicyConfigData) -and ($RetentionCompliancePolicy.Enabled -eq $true) ) {
                        $AnyPolicyEnabled =$true
                        $ConfigObject.ConfigData = "No workload covered"                                                      
                        $ConfigObject.ConfigItem = "$PolicyConfigData"   
                        $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")
                        $ConfigObjectList += $ConfigObject   
                }
                elseif(($null -ne $PolicyConfigData) -and ($RetentionCompliancePolicy.Enabled -ne $true) ) {
                    $PolicyDisabled =$true
                    $ConfigObject.ConfigData = "Policy is not enabled"                                                      
                    $ConfigObject.ConfigItem = "$PolicyConfigData"   
                    $ConfigObject.SetResult([MCCAConfigLevel]::Informational, "Pass")
                    $ConfigObjectList += $ConfigObject   
            } 
            }
            
            if (($LabelPolicy -eq $false)-and ($PolicyDisabled -eq $false)) {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>No active policy or label defined<B>"
                $ConfigObject.ConfigItem = ""
                $ConfigObject.ConfigData = "<B>Affected workloads: </B>Exchange, SharePoint, OneDrive"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
                $ConfigObjectList += $ConfigObject 
           
            }
        
            $PartialLocation=""
            foreach ($Workload in ($PartialWorkloadsStatus.Keys | Sort-Object -CaseSensitive) ) {
          
                if ($PartialWorkloadsStatus[$Workload] -eq $false) {
                    if ( $PartialLocation -eq "") {
                        $PartialLocation += "$($Workload)"
                    }else {
                        $PartialLocation += ", $($Workload)"
                    }
                }
            }
           
            if(($PartialLocation -ne "")  -and (($PolicyDisabled -eq $true) -or ($AnyPolicyEnabled -eq $true)))
            {
                $ConfigObject = [MCCACheckConfig]::new()
                $ConfigObject.Object = "<B>No policy defined for 1 or more workloads<B>"
                $ConfigObject.ConfigItem = ""
                $ConfigObject.ConfigData = "<B>Affected workloads: </B>$PartialLocation"
                $ConfigObject.SetResult([MCCAConfigLevel]::Ok, "Fail")            
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
