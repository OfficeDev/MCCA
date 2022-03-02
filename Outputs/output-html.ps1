using module "..\MCCA.psm1"

class html : MCCAOutput {

    $OutputDirectory = $null
    $DisplayReport = $True

    html() {
        $this.Name = "HTML"
    }

    RunOutput($Checks, $Collection) {
        <#

        OUTPUT GENERATION / Header

    #>

        # Obtain the tenant domain and date for the report
        $TenantDomain = ($Collection["AcceptedDomains"] | Where-Object { $_.InitialDomain -eq $True }).DomainName
        $ReportDate = "$(Get-Date -format 'dd-MMM-yyyy HH:mm') $($(Get-TimeZone).Id)"
    
        # Obtain the Remediation Report File name
    
        if ($null -eq $this.OutputDirectory) {
            $OutputDir = $this.DefaultOutputDirectory
        }
        else {
            $OutputDir = $this.OutputDirectory
        }

        $RemediationReportFileName = "$OutputDir\MCCA-$(Get-Date -Format 'yyyyMMddHHmm')-Remediation.html"
        
        # Summary
        $RecommendationCount = $($Checks | Where-Object { $_.Result -eq "Fail" }).Count
        $OKCount = $($Checks | Where-Object { $_.Result -eq "Pass" }).Count
        $InfoCount = $($Checks | Where-Object { $_.Result -eq "Recommendation" }).Count
        #>
        # Misc
        $ReportTitle = "Microsoft Compliance Configuration Analyzer"

        # Area icons
        $AreaIcon = @{}
        $AreaIcon["Default"] = "fas fa-user-cog"
        $AreaIcon["Data Loss Prevention"] = "fas fa-scroll"
    
        # Output start
        if ($null -ne $this.VersionCheck.Version) {
            $version = $($this.VersionCheck.Version.ToString())
        }
        else { 
            $Version = '' 
        }

        $output = "<!doctype html>
    <html lang='en'>
    <head>
        <!-- Required meta tags -->
        <meta charset='utf-8'>
        <meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>

        <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/all.min.css' crossorigin='anonymous'>
        <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
        <script src='https://code.jquery.com/jquery-3.3.1.slim.min.js' integrity='sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo' crossorigin='anonymous'></script>
        <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js' integrity='sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1' crossorigin='anonymous'></script>
        <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js' integrity='sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM' crossorigin='anonymous'></script>
        <script src='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/js/all.js'></script>
       

        <style>
        .navbar-custom { 
            background-color: #005494;
            color: white; 
            padding-bottom: 10px;

            
        } 
        /* Modify brand and text color */ 
          
        .navbar-custom .navbar-brand, 
        .navbar-custom .navbar-text { 
            color: white; 
            padding-top: 70px;
            padding-bottom: 10px;

        } 
        .card-header {
            background-color: #0078D4;;
            color: white; 
        }
       
        .table-borderless td,
        .table-borderless th {
            border: 0;
            padding:5px; 

        }
        .bd-callout {
            padding: 1.25rem;
            margin-top: 1.25rem;
            margin-bottom: 1.25rem;
            border: 1px solid #eee;
            border-left-width: .25rem;
            border-radius: .25rem
        }
        
        .bd-callout h4 {
            margin-top: 0;
            margin-bottom: .25rem
        }
        
        .bd-callout p:last-child {
            margin-bottom: 0
        }
        
        .bd-callout code {
            border-radius: .25rem
        }
        
        .bd-callout+.bd-callout {
            margin-top: -.25rem
        }
        
        .bd-callout-info {
            border-left-color: #5bc0de
        }
        
        .bd-callout-info h4 {
            color: #5bc0de
        }
        
        .bd-callout-warning {
            border-left-color: #f0ad4e
        }
        
        .bd-callout-warning h4 {
            color: #f0ad4e
        }
        
        .bd-callout-danger {
            border-left-color: #d9534f
        }
        
        .bd-callout-danger h4 {
            color: #d9534f
        }

        .bd-callout-success {
            border-left-color: #00bd19
        }
        .app-footer{
            background-color: #005494;
            color: white; 
            padding-top:2px; 
            padding-bottom :2px; 
        }
        .star-cb-group {
            /* remove inline-block whitespace */
            font-size: 0;
            /* flip the order so we can use the + and ~ combinators */
            unicode-bidi: bidi-override;
            direction: rtl;
            /* the hidden clearer */
          }
          .star-cb-group * {
            font-size: 1rem;
          }
          .star-cb-group > input {
            display: none;
          }
          .star-cb-group > input + label {
            /* only enough room for the star */
            display: inline-block;
            overflow: hidden;
            text-indent: 9999px;
            width: 1.7em;
            white-space: nowrap;
            cursor: pointer;
          }
          .star-cb-group > input + label:before {
            display: inline-block;
            text-indent: -9999px;
            content: ""\2606"";
            font-size: 30px;
            color: #005494;
          }
          .star-cb-group > input:checked ~ label:before, .star-cb-group > input + label:hover ~ label:before, .star-cb-group > input + label:hover:before {
            content:""\2605"";
            color: #e52;
          font-size: 30px;
            text-shadow: 0 0 1px #333;
          }
          .star-cb-group > .star-cb-clear + label {
            text-indent: -9999px;
            width: .5em;
            margin-left: -.5em;
          }
          .star-cb-group > .star-cb-clear + label:before {
            width: .5em;
          }
          .star-cb-group:hover > input + label:before {
            content: ""\2606"";
            color: #005494;
          font-size: 30px;
            text-shadow: none;
          }
          .star-cb-group:hover > input + label:hover ~ label:before, .star-cb-group:hover > input + label:hover:before {
            content: ""\2605"";
            color: #e52;
          font-size: 30px;
            text-shadow: 0 0 1px #333;
          }
        </style>

        <title>$($ReportTitle)</title>

    </head>
    <body class='app bg-light'>

        <nav class='navbar navbar-custom' >
            <div class='container-fluid'>
                <div class='col-sm' style='text-align:left'>
                    <div class='row'><div><i class='fas fa-binoculars'></i></div><div class='ml-3'><strong>Microsoft Compliance Configuration Analyzer (MCCA)</strong></div></div>
                </div>
              
                <div class='col-sm' style='text-align:right'>
                <button type='button' class='btn btn-primary' onclick='javascript:window.print();'>Print</button>
                 <BR/> 
               

                </div>
            </div>
        </nav>  
              <div class='app-body p-3'>
            <main class='main'>
                <!-- Main content here -->
                <div class='container' style='padding-top:10px;'></div>
                <div class='card'>
                        
                        <div class='card-body'>

                            <h2 class='card-title'>$($ReportTitle)</h2>"
                            
                            $Output += "<div style='text-align:right;margin-top:-65px;margin-right:8px;color:#005494;';>
				            <b>Rate this report</b>
					</div>
                         <div style='text-align:right;margin-top:-10px';>
             
                         <span class='star-cb-group'>
                            <input type='radio' id='rating-5' name='rating' value='5' onclick=""window.open('https://aka.ms/mcca-feedback-5','_blank');"" />
                            <label for='rating-5'>5</label>
                            <input type='radio' id='rating-4' name='rating' value='4' onclick=""window.open('https://aka.ms/mcca-feedback-4','_blank');"" />
                            <label for='rating-4'>4</label>
                            <input type='radio' id='rating-3' name='rating' value='3' onclick=""window.open('https://aka.ms/mcca-feedback-3','_blank');"" />
                            <label for='rating-3'>3</label>
                            <input type='radio' id='rating-2' name='rating' value='2' onclick=""window.open('https://aka.ms/mcca-feedback-2','_blank');"" />
                            <label for='rating-2'>2</label>
                            <input type='radio' id='rating-1' name='rating' value='1' onclick=""window.open('https://aka.ms/mcca-feedback-1','_blank');"" />
                            <label for='rating-1'>1</label>
                            <input type='radio' id='rating-0' name='rating' value='0' class='star-cb-clear' />
                            <label for='rating-0'>0</label>
                            </span>
                         </div>"

                            if ($(Test-Path -Path "$PSScriptRoot\..\Image\logo.jpg") -eq $True) {
                                $Output += "<img src='$PSScriptRoot\..\Image\logo.jpg' align='right' width='250px' height='150px'/>
                                "
                            }
                    
                            $Output += "<strong>Version $version </strong>
                            <p> MCCA assesses your compliance posture, highlights risks and recommends remediation steps to ensure compliance with essential data protection and regulatory standards.</p>"

                            

        
        $Output += "<table><tr><td>
                            <strong>Date</strong>  </td>
                            <td><strong>: $($ReportDate)</strong>  </td>
                            </tr>
                           
                            "
        if ($Collection["GetOrganisationConfig"] -ne "Error") {
            $OrganisationName = $Collection["GetOrganisationConfig"].DisplayName
            if (($null -ne $($OrganisationName)) -and ($($OrganisationName) -ne "")) { 

                $output += " <tr><td><strong>Organization &nbsp;</strong> </td>
                                             <td><strong>: $($OrganisationName)</strong> </td></tr>
                                             " 
            }
        }   
        if (($null -ne $($TenantDomain)) -and ($($TenantDomain) -ne "")) {
            $output += " <tr><td><strong>Tenant &nbsp;</strong> </td>
                             <td><strong>: $($TenantDomain)</strong> </td></tr>
                             " 
        }   
        $TenantGeoLocations = $Collection["GetOrganisationRegion"] | Where-Object { $_ -ne "INTL" }
        if ($TenantGeoLocations -ne "Error") {
            $RegionString = ""
            $NumberToRegionMapping = Get-NumberRegionMappingHashTable
            foreach ($Region in $TenantGeoLocations) {
                foreach ($Numbers in $($NumberToRegionMapping.Keys)) {
                    if ($($NumberToRegionMapping[$Numbers].Code) -eq $Region) {
                        if ($RegionString -eq "") {
                            $RegionString += "$($NumberToRegionMapping[$Numbers].Description)" 
                        }
                        else {
                            $RegionString += ", $($NumberToRegionMapping[$Numbers].Description)" 
                        }
                    }
                }

            }
            $output += " <tr><td><strong>Note &nbsp;</strong> </td>
                             <td><strong>:</strong>&nbsp;The following report is customized for following geolocation(s): $RegionString</td></tr>
                             " 
        }
        else {
            $output += " <tr><td><strong>Note &nbsp;</strong> </td>
                             <td><strong>:</strong>&nbsp;The following report is generalized on all geolocations</td></tr>
                             " 
        }
                            
                            
        $output += "  </table>"
        <#

                OUTPUT GENERATION / Version Warning

        #>
                                
        If ($this.VersionCheck.Updated -eq $False) {

            $Output += "
            <div class='alert alert-danger pt-2' role='alert'>
                MCCA is out of date. You're running version $($this.VersionCheck.Version) but version $($this.VersionCheck.GalleryVersion) is available! Run Update-Module MCCA to get the latest definitions!
            </div>
            
            "
        }

        $Output += "</div>
                </div>"



        <#

        OUTPUT GENERATION / Summary cards

    #>

        $Output += "<br/>"





        <#
    
        OUTPUT GENERATION / Summary

    #>

        $Output += "
    <div class='card m-3'>
    <a name='Solutionsummary'></a>

        <div class='card-header'>
          Solutions Summary
        </div>
        <div class='card-body'>"
        $Output += "<table class='table table-borderless'>
        <tr>
            <td width='20'><i class='fas fa-user-cog'></i>
            <td><strong>All Solutions</strong></td>
            <td align='right'>
                <span class='badge badge-secondary' style='padding:15px;text-align:center;width:40px;"; $output += "'>$($InfoCount)</span>
                <span class='badge badge-warning' style='padding:15px;text-align:center;width:40px;"; $output += "'>$($RecommendationCount)</span>
                <span class='badge badge-success' style='padding:15px;text-align:center;width:40px;"; $output += "'>$($OkCount)</span>
            </td>
        </tr>
        "

        ForEach ($ParentArea in ($Checks | Where-Object { $_.Completed -eq $true } | Group-Object ParentArea)) {  
            $Icon = $AreaIcon["Default"]
            If ($Null -eq $Icon) { $Icon = $AreaIcon["Default"] }
            $Output += "
        <tr >
            <td width='20'><i class='$Icon'></i>
            <td><strong>$($ParentArea.Name)</strong></td>   
        </tr>
        "    
            ForEach ($Area in ($Checks | Where-Object { $_.Completed -eq $true } | Where-Object { $_.ParentArea -eq $ParentArea.Name } | Group-Object Area)) {

                $Pass = @($Area.Group | Where-Object { $_.Result -eq "Pass" }).Count
                $Fail = @($Area.Group | Where-Object { $_.Result -eq "Fail" }).Count
                $Info = @($Area.Group | Where-Object { $_.Result -eq "Recommendation" }).Count

                $Output += 
                "
            <tr>
                <td width='20'>
                <td style='vertical-align:middle;'>&nbsp;&nbsp;<i class='fa fa-cog'></i>&nbsp;&nbsp; <a href='`#$($Area.Name)'>$($Area.Name)</a></td>
                <td align='right' style='vertical-align:middle;'>
                <span class='badge badge-secondary' style='padding:10px;text-align:center;width:30px;"; $output += "'>$($Info)</span>
                <span class='badge badge-warning' style='padding:10px;text-align:center;width:30px;"; $output += "'>$($Fail)</span>
                <span class='badge badge-success' style='padding:10px;text-align:center;width:30px;"; $output += "'>$($Pass)</span>
                </td>
            </tr>
            "
            }
        }


        $Output += "
    <tr><td colspan='3' style='text-align:right'> 
        <span class='badge badge-secondary'style='padding:5px;text-align:center'> </span>&nbsp;Recommendation
        <span class='badge badge-warning'style='padding:5px;text-align:center'> </span>&nbsp;Improvement
        <span class='badge badge-success' style='padding:5px;text-align:center'> </span>&nbsp;OK
    </td></tr></table>"
        $Output += "
        </div>
    </div>
    "

        <#

        OUTPUT GENERATION / Zones

    #>
        [bool] $UncompletedChecks = $False
        [string] $UncompletedChecksName = ""
        ForEach ($Area in ($Checks | Where-Object { $_.Completed -eq $False } | Group-Object Area)) {
            if ($UncompletedChecks -eq $False) {
                $UncompletedChecks = $True
            }
            # Each check
            if ($UncompletedChecksName -eq "") {
                $UncompletedChecksName += "Note: There was an issue in fetching $($Area.Name)"
            }
            else {
                $UncompletedChecksName += ", $($Area.Name)"
            }
        }
        if ($UncompletedChecks -eq $True) {
            $UncompletedChecksName += " information. Please try running the tool again after some time."
        
            $Output += "
        <div style='color:red;'>&nbsp;&nbsp;&nbsp;
        $UncompletedChecksName
        </div>"
        }


        ForEach ($Area in ($Checks | Where-Object { $_.Completed -eq $True } | Group-Object Area)) {

            # Write the top of the card
            $CollapseId = $($Area.Name).Replace(" " , "_")
            $Output += "<a name='$($Area.Name)'></a> 
        <div class='card m-3'>
            <div class='card-header'>
            <div class=""row"">
            <div class='col-sm' style='text-align:left; margin-top:auto; margin-bottom:auto;'><a>$($Area.Name)</a></div>
            <div class='col-sm' style='text-align:right; padding-right:10px;'> 
            <span id='more_$($CollapseId)' data-toggle='collapse' data-target='#$($CollapseId)_body'>
            <i class='fas fa-chevron-down' >&nbsp;&nbsp;</i>
            </span>
            </div>  
            </div>        
            </div>
            
            <div class='card-body collapse show' id='$($CollapseId)_body'>"

            # Each check
            [int] $count = 1 
            ForEach ($Check in ($Area.Group | Sort-Object Result -Descending)) {
                $RemediationActionsExist = $false
                $CheckCollapseId = $($CollapseId) + $count.ToString()

            
                If ($Check.Result -eq "Pass") {
                    $CalloutType = "bd-callout-success"
                    $BadgeType = "badge-success"
                    $BadgeName = "OK"
                    $Icon = "fas fa-thumbs-up"
                    $IconColor = "green"
                    $Title = $Check.PassText
                } 
                ElseIf ($Check.Result -eq "Recommendation") {
                    $CalloutType = "bd-callout-secondary"
                    $BadgeType = "badge-secondary"
                    $BadgeName = "Recommendation"
                    $Icon = "fas fa-thumbs-up"
                    $IconColor = "gray"
                    $Title = $Check.FailRecommendation
                }
                Else {
                    $CalloutType = "bd-callout-warning"
                    $BadgeType = "badge-warning"
                    $BadgeName = "Improvement"
                    $Icon = "fas fa-thumbs-down"
                    $IconColor = "#e5ad06"
                    $Title = $Check.FailRecommendation
                }

                $Output += "        
                    <div class='row border-bottom' style='padding:5px; vertical-align:middle;'>
                    <div class='col-sm-10' style='text-align:left; margin-top:auto; margin-bottom:auto;'><h6>$($Check.Name)</h6></div>
                    <div class='col' style='text-align:right;padding-right:10px;'> 
                    <h6>
                    <span class='badge $($BadgeType)'>$($BadgeName)</span>&nbsp;&nbsp;
                    <i class='fas fa-chevron-down' data-toggle='collapse' data-target='#$($CheckCollapseId)'></i>
                    </h6>
                    </div>  
                    </div> "
                $Output += "  
                    <div class='row collapse' id='$($CheckCollapseId)'>
                        <div class='bd-callout $($CalloutType) b-t-1 b-r-1 b-b-1 p-3' >
                            <div class='container-fluid'>
                                <div class='row'>
                                    <div><i class='$($Icon)' color='$($IconColor)'></i></div>
                                    <div class='col-8'><h6>$($Title)</h6></div>
                                   
                                </div>"

                if ($Check.Importance) {

                    $Output += "
                                <div class='row p-3'>
                                    <div><p>$($Check.Importance)</p></div>
                                </div>"

                }
                        
                        
                If ($Check.ExpandResults -eq $True) {
                             

                    # We should expand the results by showing a table of Config Data and Items
                    $Output += "
                            <div class='row pl-2 pt-3'>"
                    if ($Check.Control -ne "Compliance Manager") {
                        $Output += "  <table class='table'>
                                    <thead class='border-bottom'>
                                        <tr>"

                        If ($Check.CheckType -eq [CheckType]::ObjectPropertyValue) {
                            # Object, property, value checks need three columns
                            $Output += "
                                <th align='center' text-align='center'>$($Check.ObjectType)</th>
                                <th align='center' text-align='center'>$($Check.ItemName)</th>
                                <th align='center' text-align='center'> $($Check.DataType)</th>
                                <th align='center' text-align='center'>Status</th>
                                "    
                        }
                        Else {
                            $Output += "
                                <th  align='center' text-align='center'>$($Check.ItemName)</th>
                                <th align='center' text-align='center'>$($Check.DataType)</th>
                                <th align='center' text-align='center'>Status</th>
                                "     
                        }

                        $Output += "
                                            <th style='width:50px'></th>
                                        </tr>
                                    </thead>
                                    <tbody>
                            "

                        ForEach ($o in $Check.Config | Sort-Object Level -Descending) {
                            $ActionRequired = $false
                            if ($o.Level -ne [MCCAConfigLevel]::None -and $o.Level -ne [MCCAConfigLevel]::Recommendation) {
                                $oicon = "fas fa-check-circle text-success"
                                $LevelText = $o.Level.ToString()
                            }
                            ElseIf ($o.Level -eq [MCCAConfigLevel]::Recommendation) {
                                $oicon = "fas fa-info-circle text-muted"
                                $LevelText = $o.Level.ToString()
                            }
                            Else {
                                $oicon = "fas fa-times-circle text-danger"
                                $LevelText = "Improvement"
                                $ActionRequired = $true 
                            }

                            $Output += "
                                <tr>
                                "
                            if ($($o.RemediationAction)) {
                                $RemediationActionsExist = $true
                            }
                            If ($Check.CheckType -eq [CheckType]::ObjectPropertyValue) {
                                # Object, property, value checks need three columns
                                $Output += "
                                        <td>$($o.Object)</td>
                                        <td style='word-wrap:break-word;' width = '35%'>$($o.ConfigItem)</td>
                                        <td style='word-wrap:break-word;' width = '30%'>$($o.ConfigData)</td>
                                    "
                            }
                            Else {
                                $Output += "
                                        <td>$($o.ConfigItem)</td>
                                        <td style='word-wrap:break-word;' width = '35%'>$($o.ConfigData)</td>
                                    "
                            }

                            $Output += "
                                    <td style='text-align:left'>
                                        <div class='row badge badge-pill badge-light'>"
                            if ($o.Level -eq [MCCAConfigLevel]::Informational) {
                                $Output += "<span style='vertical-align: left;'>$($LevelText)</span><br/></div>"  
                            }
                            else {
                                $Output += "<span class='$($oicon)' style='vertical-align: left;'></span>
                                            <span style='vertical-align: left;'>$($LevelText)</span><br/></div>"
                            }
                            if ($ActionRequired -eq $true -and $($o.RemediationAction)) {
                                $Output += " <span style='vertical-align: left;'><small><center>Remediation Available</center></small></span> "
                            }
                            $Output += " 
                                    </td>
                                </tr>
                                "

                            # Recommendation segment
                            #if($o.Level -eq [MCCAConfigLevel]::Recommendation)
                            #{
                            if (($null -ne $($o.InfoText)) -and ($($o.InfoText) -ne "" ) ) {
                                        
                                $Output += "
                                    <tr>"
                                If ($Check.CheckType -eq [CheckType]::ObjectPropertyValue) {
                                    $Output += "<td colspan='4' style='border: 0;'>"
                                }
                                else {
                                    $Output += "<td colspan='3' style='border: 0;'>"
                                }
                                   
                                $Output += "
                                    <div class='alert alert-light' role='alert' style='text-align: left;'>
                                    <span class='fas fa-info-circle text-muted' style='vertical-align: left; padding-right:5px'></span>
                                    <span style='vertical-align: middle;'>$($o.InfoText)</span>
                                    </div>
                                    "
                                    
                                $Output += "</td></tr>
                                    
                                    "
                            }
                                    
                        }

                        #}

                        $Output += "
                                    </tbody>
                                </table>"
                    }
                    # If any links exist
                    If ($Check.Links) {
                        $Output += "
                                <table class='table'> <tr>"                                 
                        $LinksInfo = $Check.Links
                        [int] $CountOfLinks = $LinksInfo.Keys.Count
                        [int] $itr = 0
                        $LinksNameValuePair = $LinksInfo.GetEnumerator() | Sort-Object -Property Name
                        while ($itr -lt $CountOfLinks) {
                            $Output += "

                                   
                                    <td style='padding-top:20px;'><i class='fas fa-external-link-square-alt'></i>&nbsp;<a href='$($LinksNameValuePair.Value[$itr])' target=""blank"">$($LinksNameValuePair.Name[$itr])</a></td>
                                    
                                    "
                            $itr = $itr + 1
                        }

                        if ($RemediationActionsExist -eq $true) {
                                    
                            $Output += "
                            
                                    <td ><a class='btn btn-primary' href='$($RemediationReportFileName)' target='_blank' role='button'>Remediation Script</a></td>
                                     
                                    "
                                    
                        }
                        $Output += "
                               </tr> </table>
                                "
                        $Output += "
                                </table>
                                "
                    }

                    $Output += "
                            </div>"

                }
                        

                $Output += "
                            </div>
                        </div> </div> "
                $count += 1
            }            

            # End the card
            $Output += "   <div class='col-sm' style='text-align:right; padding-right:10px;'>  <a href='#Solutionsummary'>Go to Solutions Summary</a></div>
            </div>
                      

        </div>"
        }
        <#

        OUTPUT GENERATION / Footer

    #>

        $Output += "
            </main>
            <center>Bugs? Issues? Suggestions? <a href='https://github.com/OfficeDev/MCCA'>GitHub</center>
            </div>
            <footer class='app-footer'>
            <p><center><i>&nbsp;&nbsp;&nbsp;&nbsp;Disclaimer: Recommendations from  (MCCA) should not be interpreted as a guarantee of compliance. It is up to you to evaluate and validate the effectiveness of customer controls per your regulatory environment. <br>
               </i></center> </p></footer>
        </body>
    </html>"


        # Write to file

       
        $Tenant = $(($Collection["AcceptedDomains"] | Where-Object { $_.InitialDomain -eq $True }).DomainName -split '\.')[0]
        $ReportFileName = "MCCA-$($tenant)-$(Get-Date -Format 'yyyyMMddHHmm').html"

        $OutputFile = "$OutputDir\$ReportFileName"

        $Output | Out-File -FilePath $OutputFile

        If ($this.DisplayReport) {
            Invoke-Expression $OutputFile
        }

        $this.Completed = $True
        $this.Result = $OutputFile

    }

}