# Overview

Microsoft Compliance Configuration Analyzer (MCCA) is a tool which, on execution, generates a report highlighting known issues in your compliance configurations in achieving data protection guidelines and recommends best practices to follow.

# What is Microsoft Compliance Configuration Analyzer (MCCA)?

It is a PowerShell-based utility that will fetch your tenant’s current configurations & validate these configurations against Microsoft 365 recommended best practices. These best practices are based on a set of controls that include key regulations and standards for data protection and general data governance. MCCA then provides you with an actionable status report for improving your compliance posture.

# Why should I use it?

Often tenants face challenges in diagnosing their compliance posture & ensuring that they have the right configurations in place to protect their environment completely. These are largely manual processes which tend to be time consuming & allow for human error. Furthermore, with the evolving compliance landscape the risk of blind spots also increases.
MCCA is a diagnostic tool that will report the status of your current configurations. This allows you to focus efforts more on making the right configurations. 

# What is in scope?

This version will provide you recommendations for the M365 Compliance solutions listed below. We will keep adding more solutions & richer recommendations in future versions of this tool.
  
        1.	Microsoft Information Protection
            a. 	Data Loss Prevention
            b.	Information Protection
        2.	Microsoft Information Governance
            a.	Information Governance
            b.	Records Management
        3.	Insider Risk
            a.	Communication Compliance
            b.	Insider Risk Management
        4.	Discovery & Response
            a.	Audit
            b.	eDiscovery

# That is awesome! How do I run it?

#   Pre-Requisites

Before running the tool, you should confirm your Microsoft 365 subscription and any add-ons. To access and use MCCA, your organization must have one of the following subscriptions or add-ons:
   
    •	Microsoft 365 E5 subscription (paid or trial version)
    •	Microsoft 365 E3 subscription + the Microsoft 365 E5 Compliance add-on

You will be able to run this tool without an E5 subscription or M365 E5 Compliance add-on, but MCCA will still report statuses for E5 workloads & capabilities.

For running the tool:
     
     1.	You must have PowerShell version 5.1 or above to run this tool.
     2.	You must have Exchange Online PowerShell module (You can follow either of the following 2 methods to download the same)
        •	Exchange Online PowerShell V2 module (v2.0.3 or higher) that is available via the PowerShell gallery:
                Install-Module -Name ExchangeOnlineManagement (v2.0.3 or higher)
     3.	You must have appropriate role/user permissions to be able to run this tool. Refer to the ReadMe.docx
  Other roles within the organisation (not listed in the table) may not be able to run the tool or they may be able to run the tool with limited information in the final report.

# Install Guide	

Step 1: Open PowerShell in administrator mode
    
Step 2: Install MCCA 
   
    Install-Module -Name MCCAPreview

Step 3: Generate MCCA Report
  
    Use the following cmdlet to generate the MCCA report.
    Get-MCCAReport
    
   This will generate a report based on the geolocation of your tenant. If an error occurs while fetching your tenant’s geolocation, you will get a report covering all supported geolocations.
    
 You can learn more about this cmdlet by running the following.
    
    Get-Help Get-MCCAReport

  Input Parameters	
   You can also get a tailored report based on specific input parameters listed below.

   1.	Geolocation
         
     Get-MCCAReport -Geo @(1,7)
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

    Note: As an add-on, the report will always include MCCA supported international sensitive information types like SWIFT Code, Credit Card Number etc.

   2.	Solutions
          
    Get-MCCAReport -Solution @(1,7)
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

   3.	Multiple Parameters
            
          Get-MCCAReport -Solution @(1,7) -Geo @(9)
          
         This will generate a report only on for the solutions entered by you and based on the regions you have selected. 
  In either of the cases, there will be a prompt to enter your credentials. Once you enter your credentials, MCCA will run for a while and an HTML report will be generated.

# License
We use the following open source components in order to generate the report:
    •	Bootstrap, MIT License - https://getbootstrap.com/docs/4.0/about/license/
    •	Fontawesome, CC BY 4.0 License - https://fontawesome.com/license/free
    •	clipboard.js v1.5.3, MIT License - https://cdn.jsdelivr.net/clipboard.js/1.5.3/clipboard.min.js


# Frequently Asked Questions (FAQ)
 Please refer to Readme.docx

#	This tool is awesome! How do I provide feedback and suggestions for future versions?
Please share your feedback & suggestions with us using this [form](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR-ItstQd6pNMqw0W9LKA5vxUOFNGUFgxRDJFTkg3VE5NQTQwTUVVVDNVMi4u). We are dying to hear from you. :)

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

# Trademark

Trademarks This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow Microsoft's Trademark & Brand Guidelines. Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.
