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
        if($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovGCCHigh")
        {
            $this.Links = @{
                "Visit Compliance Manager"              = "https://aka.ms/mcca-gcch-cm-compliance-manager"
                "Learn more about Compliance Manager"    = "https://aka.ms/mcca-cm-docs-learn-more"
                "Compliance Manager Quickstart Guide" = "https://aka.ms/mcca-cm-docs-action"
           
            }
        }elseif ($this.ExchangeEnvironmentNameForCheck -ieq "O365USGovDoD") 
        {
            $this.Links = @{
                "Visit Compliance Manager"              = "https://aka.ms/mcca-dod-cm-compliance-manager"
                "Learn more about Compliance Manager"  = "https://aka.ms/mcca-cm-docs-learn-more"
                "Compliance Manager Quickstart Guide" = "https://aka.ms/mcca-cm-docs-action"
           
            }
        }else
        {
        $this.Links = @{
            "Visit Compliance Manager"              = "https://aka.ms/mcca-cm-compliance-manager"
            "Learn more about Compliance Manager"   = "https://aka.ms/mcca-cm-docs-learn-more"
            "Compliance Manager Quickstart Guide" = "https://aka.ms/mcca-cm-docs-action"
       
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

