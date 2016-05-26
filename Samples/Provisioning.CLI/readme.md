# Provision sites from command line #

### Summary ###
This sample demonstrates how to extract and apply templates using the CSOM model.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Prerequisites ###
None

### Solution ###
Solution | Author(s)
---------|----------
Provisioning.CLI | Konrad Brunner (**Trivadis**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
0.1  | May 25th 2016 | Initial release, able to export a template
0.2  | May 26th 2016 | Added functionality to apply templates

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Key components of the sample #

The sample console application contains the following:

- **Provisioning Console project**, which contains:
    - **Program.cs**: Contains the code
    - **TokenHelper.cs**: Contains the helper code that enables the application to get the required permissions for creating site collections. This file is not a default component of the console application template in Visual Studio 2013. You can get this file by creating a provider-hosted or autohosted add-in for SharePoint and copying the file from the remote web application to the console app project.

# Configure the sample #
There is no special configuration required in this project.

# Build, deploy, and run the sample #
## To build and deploy the console application ##

1. Build the solution in Visual Studio
2. Run the command line tool

## To run the sample ##

### Extracting a template ###

Provisioning.CLI.Console.exe -Action ExtractTemplate -Url https://yourdomain.sharepoint.com/sites/yoursitecollection/yourweb -LoginMethod SPO -User youruser@yourdomain.onmicrosoft.com -OutFile "C:\template.xml" -Password *******

### Apply a template ###

Provisioning.CLI.Console.exe -Action ApplyTemplate -Url https://yourdomain.sharepoint.com/sites/yoursitecollection/yourweb -LoginMethod SPO -User youruser@yourdomain.onmicrosoft.com -InFile "C:\template.xml" -Password *******

### Apply mutiple templates with absolute paths in file ###

Provisioning.CLI.Console.exe -Action ApplyTemplate -LoginMethod SPO -User youruser@yourdomain.onmicrosoft.com -InFile "C:\sitesAbsPath.xml" -Password *******

### Apply mutiple templates with relative paths in file ###

Provisioning.CLI.Console.exe -Action ApplyTemplate -LoginMethod SPO -User youruser@yourdomain.onmicrosoft.com -InFile "C:\sitesRelPath.xml" -Url https://yourdomain.sharepoint.com/sites/yoursitecollection/yourweb -Password *******





