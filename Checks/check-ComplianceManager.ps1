using module "..\MCCA.psm1"

class ComplianceManager : MCCACheck {
    <#
    

    #>

    ComplianceManager() {
        $this.Control = "Compliance Manager"
        $this.ParentArea = "Compliance Manager"
        $this.Area = "Compliance Manager"
        $this.Name = "Use Compliance Manager to manage your compliance posture"
        $this.PassText = "Your organization should use Compliance Manager to manage your compliance posture"
        $this.FailRecommendation = "Your organization should use Compliance Manager to manage your compliance posture"
        $this.Importance = "Compliance Manager is an end-to-end solution in the Microsoft 365 compliance center for managing and tracking compliance activities.  It simplifies compliance and helps reduce risk. Compliance Manager translates complex regulatory requirements to specific controls and through compliance score, provides a quantifiable measure of compliance. It offers intuitive compliance management, a vast library of scalable assessments, and built-in automation.
        Its a great place to begin your compliance journey because it gives you an initial assessment of your compliance posture the first time you visit."
        $this.ExpandResults = $true
        if($this.ExchangeEnvironmentNameForCheck -eq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Visit Compliance Manager"              = "https://compliance.microsoft.us/compliancescore?filter=%7B%22Solution%22:%5B%22Compliance%20Score%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
                "Learn more about Compliance Manager"                       = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager?view=o365-worldwide"
                "Compliance Manager Quickstart Guide" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager-quickstart?view=o365-worldwide"
           
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -eq "O365USGovDoD") 
        {
            $this.Links = @{
                "Visit Compliance Manager"              = "https://compliance.apps.mil/compliancescore?filter=%7B%22Solution%22:%5B%22Compliance%20Score%22%5D,%22Status%22:%5B%22None%22,%22NotAssessed%22,%22Passed%22,%22FailedLowRisk%22,%22FailedMediumRisk%22,%22FailedHighRisk%22,%22NotInScope%22,%22ToBeDetermined%22,%22CouldNotBeDetermined%22,%22PartiallyTested%22,%22Select%22%5D%7D&viewid=ImprovementActions"
                "Learn more about Compliance Manager"                       = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager?view=o365-worldwide"
                "Compliance Manager Quickstart Guide" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager-quickstart?view=o365-worldwide"
           
            }
        }else
        {
        $this.Links = @{
            "Visit Compliance Manager"              = "https://compliance.microsoft.com/compliancemanager"
            "Learn more about Compliance Manager"                       = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager?view=o365-worldwide"
            "Compliance Manager Quickstart Guide" = "https://docs.microsoft.com/en-us/microsoft-365/compliance/compliance-manager-quickstart?view=o365-worldwide"
       
        }
        }
    }

    <#
    
        RESULTS
    
    #>

    GetResults($Config) {   
        
            $ConfigObjectList = @()
            
            $ConfigObject = [MCCACheckConfig]::new()
           
            $ConfigObject.SetResult([MCCAConfigLevel]::Recommendation, "Pass")
            $this.AddConfig($ConfigObject)
            $ConfigObjectList += $ConfigObject
            
            $this.Completed = $True
        }
        
    }

