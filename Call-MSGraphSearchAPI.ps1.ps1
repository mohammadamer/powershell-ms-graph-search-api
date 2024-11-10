
<#
 .SYNOPSIS
    Calls Microsoft Graph Search API to search SharePoint content for specific keywords.

 .DESCRIPTION
   Calls Microsoft Graph Search API to search SharePoint content for specific keywords.

    How it works?
    You can customize the entityTypes in line 47 as per your requirements.
    entityTypes = @("driveItem", "listItem")

    You can customize the search query string in line 75 as per your requirements.
     query       = @{ queryString = "search-query" }

    Changing the search query to whatever you want to search for in the SharePoint content."

 .INPUTS
    N/A

 .OUTPUTS
    N/A


.PARAMETER ClientId
    Azure AD App Registraion client ID of SharePointOnline app registration
 .PARAMETER CertificatePassword
    Azure AD App Registration certificate password of the SharePointOnline app registration in a secure string format.
 .PARAMETER CertificatePath
    Azure AD App Registration certificate path of the SharePointOnline app registration.
 .PARAMETER Tenant
    Microsoft 365 tenant name.

 .EXAMPLE
     .\Call-MSGraphSearchAPI.ps1 -SiteUrl https://contoso.sharepoint.com -ClientId xxxxxxxx-xxxxx-xxxx-xxxx-xxxxx -CertificatePath "C:\cert.pfx" -CertificatePassword (ConvertTo-SecureString "abc" -AsPlainText -Force)  -Tenant "contoso.onmicrosoft.com"
 .LINK
    N/A
#>
param (
   [Parameter(Mandatory = $true)][System.String]$ClientId,
   [Parameter(Mandatory = $true)][System.Security.SecureString]$CertificatePassword = (ConvertTo-SecureString $env:SharePointOnlineCertificatePassword -AsPlainText -Force),
   [Parameter(Mandatory = $true)][System.String]$CertificatePath,
   [Parameter(Mandatory = $true)][System.String]$Tenant,
   [Parameter(Mandatory = $true)][System.String]$SiteUrl

)
Set-StrictMode -Version Latest
$DebugPreference = "SilentlyContinue"
$reportName = "results.csv"

Connect-PnPOnline -Url $SiteUrl -ClientID $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Tenant $Tenant
$accessToken = Get-PnPGraphAccessToken
$uri = 'https://graph.microsoft.com/v1.0/search/query'

$headers = @{
   "Authorization" = "Bearer $($accessToken)"
   "Content-Type"  = "application/json"
}

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
$searchQueryADS | ConvertTo-Json -Depth 10

$response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body ($searchQueryADS | ConvertTo-Json -Depth 10)
$jsonResponse = $response.value | ConvertTo-Json -Depth 10

# Convert JSON response to PowerShell object
$responseObject = ConvertFrom-Json -InputObject $jsonResponse

# Inspect the top-level properties
$responseObject

# If 'hitsContainers' doesn't show up, try accessing other properties
$responseObject.hitsContainers

# If hitsContainers is an array, inspect the first element
$responseObject.hitsContainers[0]

# Assuming $jsonResponse contains the correct JSON
$responseObject = ConvertFrom-Json -InputObject $jsonResponse

# Check if hitsContainers exist
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
        
         # Conditionally add 'Name' property if it exists
         if ($hit.resource.PSObject.Properties['name']) {
            $properties['Name'] = $($hit.resource.name)
         }
        
         # Create PSObject and export to CSV
         New-Object PSObject -Property $properties | Export-Csv -Path $reportName -NoTypeInformation -Append -Encoding UTF8
      }
   }
}
else {
   Write-Host "hitsContainers not found. Here is the full response:"
   $responseObject | ConvertTo-Json -Depth 10
}