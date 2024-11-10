# Using PowerShell to Call the Microsoft Graph Search API and Export Results to CSV


### Introduction
Microsoft Graph offers a powerful way to interact with Microsoft 365 data. Using the Microsoft Graph Search API, you can retrieve content from SharePoint and other sources in Microsoft 365 by running customized search queries. In this article, we'll walk through a PowerShell script that calls the Microsoft Graph Search API, searches SharePoint content based on specified keywords, and exports the results into a CSV file for easy analysis.


### Prerequisites
Before running the script, ensure that the following prerequisites are met:

1. Azure AD App Registration: Register an application in Azure Active Directory with the necessary Microsoft Graph Application Permissions. These permissions are required to access Microsoft 365 data through the API:
   * Sites.Read.All: Allows the application to read items in all site collections.
   * Files.Read.All: Allows the application to read all files the signed-in user can access.
   * People.Read: Allows the application to read information about people relevant to the user.
   * Calendars.Read: Allows the application to read calendars the user has access to.
2. PnP PowerShell Module: Install this by running Install-Module -Name PnP.PowerShell if not already installed.
3. Certificate for Authentication: A certificate (.pfx) that will be used to authenticate with Azure AD.

#### Script Overview
This PowerShell script takes advantage of Microsoft Graph and PnP PowerShell modules to connect to SharePoint Online, authenticate, and execute search queries. We’ll go over each section of the script to show how it works, and we’ll demonstrate how to customize it to meet specific search needs.

### Script Parameters
The script accepts the following parameters:
* ClientId: The Azure AD App Registration client ID for the SharePoint Online application.
* CertificatePassword: The password for the Azure AD App Registration certificate, used for secure authentication.
* CertificatePath: The file path to the certificate (.pfx) used for Azure AD App Registration.
* Tenant: The Microsoft 365 tenant name (e.g., contoso.onmicrosoft.com).
* SiteUrl: The URL of the SharePoint site you wish to search.


### Code Overview
```
<#
 .SYNOPSIS
    Calls Microsoft Graph Search API to search SharePoint content for specific keywords.
 .DESCRIPTION
    This script queries the Microsoft Graph Search API to find content in SharePoint based on a keyword search.
 .PARAMETER ClientId
    Azure AD App Registration client ID for SharePoint Online.
 .PARAMETER CertificatePassword
    Certificate password for Azure AD App Registration.
 .PARAMETER CertificatePath
    Certificate path for Azure AD App Registration.
 .PARAMETER Tenant
    Microsoft 365 tenant name.
 .EXAMPLE
     .\Call-MSGraphSearchAPI.ps1 -SiteUrl https://contoso.sharepoint.com -ClientId <ClientId> -CertificatePath <Path> -CertificatePassword <Password> -Tenant <Tenant>
#>
param (
   [Parameter(Mandatory = $true)][System.String]$ClientId,
   [Parameter(Mandatory = $true)][System.Security.SecureString]$CertificatePassword,
   [Parameter(Mandatory = $true)][System.String]$CertificatePath,
   [Parameter(Mandatory = $true)][System.String]$Tenant,
   [Parameter(Mandatory = $true)][System.String]$SiteUrl
)
```

### Setp 1 : Authenticate and Set Up Access
The script uses the Connect-PnPOnline cmdlet to authenticate with the provided credentials and access token, making a connection to SharePoint.

```
Connect-PnPOnline -Url $SiteUrl -ClientID $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Tenant $Tenant
$accessToken = Get-PnPGraphAccessToken
$uri = 'https://graph.microsoft.com/v1.0/search/query'
```

### Step 2: Define Search Query and Headers
The $searchQueryADS variable contains the search parameters, specifying the types of content to retrieve and any additional fields of interest, such as file name, URL, and timestamps. In this example, we're querying for driveItem and listItem entity types, which represent files and list items in SharePoint.

```
$searchQueryADS = @{
   requests = @(
      @{
         entityTypes = @("driveItem", "listItem")
         query       = @{ queryString = "viva connection" }
         fields      = @("id", "name", "createdDateTime", "lastModifiedBy", "lastModifiedDateTime", "webUrl")
         region      = "EMEA"
      }
   )
}
```


### Step 3: Send the Search Request
Using Invoke-RestMethod, we send a POST request with the search query to the Microsoft Graph Search API. This returns a JSON response containing the search results.

```
$response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body ($searchQueryADS | ConvertTo-Json -Depth 10)
$jsonResponse = $response.value | ConvertTo-Json -Depth 10
$responseObject = ConvertFrom-Json -InputObject $jsonResponse
```


### Step 4: Extract and Export Data
If the response contains data, the script iterates through the hitsContainers array, extracting information like file name, URL, and timestamps. It then exports these details to a CSV file for easy review.

```
if ($responseObject.hitsContainers) {
   $hitsContainers = $responseObject.hitsContainers
   foreach ($container in $hitsContainers) {
      $hits = $container.hits
      foreach ($hit in $hits) {
         $properties = [ordered]@{
            Name         = ''
            lastModified = $($hit.resource.lastModifiedDateTime)
            ModifiedBy   = $($hit.resource.lastModifiedBy.user.displayName) 
            Url          = $($hit.resource.webUrl) 
            Created      = $($hit.resource.createdDateTime)
            Id           = $($hit.resource.id)
         }
         if ($hit.resource.PSObject.Properties['name']) {
            $properties['Name'] = $($hit.resource.name)
         }
         New-Object PSObject -Property $properties | Export-Csv -Path "Report.csv" -NoTypeInformation -Append -Encoding UTF8
      }
   }
}
else {
   Write-Host "hitsContainers not found. Here is the full response:"
   $responseObject | ConvertTo-Json -Depth 10
}
```

### Step 5: Run the Script
Execute the script with the required parameters. For example:


```
.\Call-MSGraphSearchAPI.ps1 -SiteUrl https://contoso.sharepoint.com -ClientId xxxxxxxx-xxxxx-xxxx-xxxx-xxxxx -CertificatePath "C:\cert.pfx" -CertificatePassword (ConvertTo-SecureString "yourpassword" -AsPlainText -Force) -Tenant "contoso.onmicrosoft.com"
```

### Conclusion
With this PowerShell script, you can automate the process of searching SharePoint content using Microsoft Graph, retrieving specific fields, and exporting the results into a CSV file for easy reporting. This script provides a flexible framework for using Microsoft Graph API to meet your search needs across SharePoint and beyond. You can further customize it by adjusting the query string, entity types, and output fields to fit your specific use case.