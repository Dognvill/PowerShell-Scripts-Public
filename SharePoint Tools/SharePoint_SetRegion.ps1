<#
.SYNOPSIS
This script sets the regional settings (locale and time zone) for all SharePoint Online web sites in a given tenant.

.DESCRIPTION
The script connects to the SharePoint Online Admin Center and retrieves all site collections,
excluding specific templates like Search Center, MySite Host, App Catalog, Content Type Hub, eDiscovery, and Bot Sites.

For each site, it sets the regional settings to the provided locale ID and time zone ID.
It includes all subsites and the root site within each site collection.

.PARAMETERS
$TenantAdminURL: The URL of the SharePoint Online Admin Center.
$LocaleId: The locale ID to set for the web sites (3081 for Australia).
$TimeZoneId: The time zone ID to set for the web sites (76 for Australia).

.NOTES
- This will run the script with the parameters specified in the script itself.
- Requires SharePoint PnP PowerShell module installed.
- Requires admin permissions to access and update the SharePoint sites.
- ExecutionPolicy should be set to allow script execution.

.AUTHOR
John Bignold
#>

# Parameter
$TenantAdminURL = "https://dogsnville-admin.sharepoint.com"
$LocaleId = 3081 # AUS
$TimeZoneId = 76 # Australia
 
#Function to Set Regional Settings on SharePoint Online Web
Function Set-RegionalSettings
{ 
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $true, ValueFromPipeline = $True)] $Web
    )
  
    Try {
        Write-host -f Yellow "Setting Regional Settings for:"$Web.Url
        #Get the Timezone
        $TimeZone = $Web.RegionalSettings.TimeZones | Where-Object {$_.Id -eq $TimeZoneId} 
        #Update Regional Settings
        $Web.RegionalSettings.TimeZone = $TimeZone
        $Web.RegionalSettings.LocaleId = $LocaleId
        $Web.Update()
        Invoke-PnPQuery
        Write-host -f Green "`tRegional Settings Updated for "$Web.Url
    }
    Catch {
        write-host "`tError Setting Regional Settings: $($_.Exception.Message)" -foregroundcolor Red
    }
}
 
#Connect to Admin Center
Connect-PnPOnline -Url $TenantAdminURL -Interactive
   
#Get All Site collections - Exclude: Seach Center, Mysite Host, App Catalog, Content Type Hub, eDiscovery and Bot Sites
$SitesCollections = Get-PnPTenantSite | Where -Property Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
   
#Loop through each site collection
ForEach($Site in $SitesCollections)
{
    #Connect to site collection
    Connect-PnPOnline -Url $Site.Url -Interactive
  
    #Call the Function for all webs
    Get-PnPSubWeb -Recurse -IncludeRootWeb -Includes RegionalSettings, RegionalSettings.TimeZones | ForEach-Object { Set-RegionalSettings $_ }
}