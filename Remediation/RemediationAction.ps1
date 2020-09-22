using module "..\MCCA.psm1"

class Action : RemediationAction
{

    $OutputDirectory=$null
    $DisplayReport=$True

    Action()
    {
        $this.Name="Action"
    }

    RunOutput($Checks,$Collection)
    {
    <#

        OUTPUT GENERATION / Header

    #>

    # Obtain the tenant domain and date for the report
    #$TenantDomain = ($Collection["AcceptedDomains"] | Where-Object {$_.InitialDomain -eq $True}).DomainName

    # Misc
    $ReportTitle = "Microsoft Compliance Configuration Analyzer Remediation Report"

    
    

    # Output start
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
        <script src='https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js'></script>
        <script src='https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js'></script>
        <style>
        .w3-code, .w3-codespan,#w3-exerciseform .exerciseprecontainer 
        .w3-code {font-size:14px;}
        .w3-codespan {font-size:15px;}
        
        .navbar-custom { 
            background-color: #005494;
            color: white; 
            padding-bottom: 10px;
        }
        .card-header {
            background-color: #0078D4;
            color: white; 
            }   
         
        .codetable{
            background-color:#FFFFFF;
            border: 1px solid #d9d9d9;
            border-collapse: collapse;
            width:90%;
            margin-left:auto;margin-right:auto;

            }
        tr {
            border: solid #d9d9d9;
            border-width: 1px 0;
        }
        tr:first-child {
            border-top: none;
        }
        tr:last-child {
            border-bottom: none;
        }
        .btn {
            background: #d9d9d9;
            outline: none !important;
            box-shadow: none !important;
        }
        .btn:focus {
            background: #009900; 
        
        }
        </style>
        <title>$($ReportTitle)</title>

    </head>
    <body class='app header-fixed bg-light'>
           <table class='table'>
           <nav class='navbar fixed-top navbar-custom'>
           <div class='container-fluid'>
               <div class='col-sm' style='text-align:left'>
                   <div class='row'><div><i class='fas fa-binoculars'></i></div><div class='ml-3'><strong>MCCA</strong></div></div>
               </div>
               
               <div class='col-sm' style='text-align:right'>
                  
               </div>
           </div>
       </nav>  
       <div class='app-body p-3'>
       <br/><br/>   <br/>   
                  <h2 class='card-title'>&nbsp;$($ReportTitle)</h2>           
                 <p>&nbspThis report details recommended remediations based on your MCCA report.</p> 
        </div></table> "
    <#
    
        OUTPUT GENERATION / Summary

    #>
       $Output +=  "<div class=""w3-code notranslate htmlHigh"">"
       [int] $count = 1
    ForEach ($Area in ($Checks | Where-Object {$_.MCCARemediationInfo.RemediationAvailable -eq $True}|  Group-Object Area ))
    {
         $Output += 
          "
          <div class='card m-3'>
          <div class='card-header'>
          $($Area.Name)</div>
         
         <div class='card-body'>
         "                            
        # Each check
        ForEach ($Check in ($Area.Group | Sort-Object Result -Descending)) 
        {        
            ForEach($o in $Check.Config)
             {                  
                if($($o.RemediationAction))
                 {
                    # div identifier for each remediation script                
                    $divid = "div"+ $count.ToString() 
                    $Output += "
                    
                    <h5><B>$($Check.Name)</B></h5>   
                    &nbsp;&nbsp;&nbsp;$($Check.MCCARemediationInfo.RemediationText)<br><br>"
                    

                    # We should expand the results by showing a table of Config Data and Items
                     $Output +=  " <table class='codetable'>
                     <tr style=""background-color: #e6e6e6;""><td>
                     <div class=""container-fluid"">
                     <div class=""row"">
                     <div class='col-sm' style='text-align:left; margin-top:auto; margin-bottom:auto;'>PowerShell</div>
                     <div class='col-sm' style='text-align:right; padding-right:0;'>
                     <button type = ""button"" class = ""btn btn-sm active rounded-0"" data-clipboard-action='copy' data-clipboard-target='div#$divid'><i class=""fa fa-clipboard"" aria-hidden=""true""></i>&nbsp; Copy</button>
                     </div>
                     </div>
                     </div>
                     </td></tr>"
                    $Output += "<tr style=""background-color: #f2f2f2;""><td><br><div id=""$divid"">$($o.RemediationAction)</div><br></td></tr>"
                    $Output +=  " </table><br>"
                 }
                     # Object, property, value checks need three columns             
             }
             $count = $count + 1
        }           

       $Output +=  "</div></div>"

    }
        $Output +=  "</div>"
    <#

        OUTPUT GENERATION / Footer

    #>

    $Output += "
           

            <footer class='app-footer'>
            <footer class='app-footer'>
                <p><center><i>&nbsp;&nbsp;&nbsp;&nbsp;Disclaimer: Recommendations from MCCA should not be interpreted as a guarantee of compliance. 
                It is up to you to evaluate and validate the effectiveness of customer controls per your regulatory environment.<br>
               </i></center> </p>
            </footer>
            </footer>
            <script>
                var clipboard = new Clipboard('.btn');
                clipboard.on('success', function(e) {
                  console.log(e);
                });
                clipboard.on('error', function(e) {
                  console.log(e);
                });
            </script>
        </body>
    </html>"


        # Write to file

        if($null -eq $this.OutputDirectory)
        {
            $OutputDir = $this.DefaultOutputDirectory
        }
        else 
        {
            $OutputDir = $this.OutputDirectory
        }

        #$Tenant = $(($Collection["AcceptedDomains"] | Where-Object {$_.InitialDomain -eq $True}).DomainName -split '\.')[0]
        $ReportFileName = "MCCA-$(Get-Date -Format 'yyyyMMddHHmm')-Remediation.html"

        $OutputFile = "$OutputDir\$ReportFileName"

        $Output | Out-File -FilePath $OutputFile
        $this.Completed = $True
        $this.Result = $OutputFile

    }

}