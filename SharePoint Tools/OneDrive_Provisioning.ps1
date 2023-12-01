<#
.SYNOPSIS
A PowerShell script to automate the provisioning of OneDrive personal sites for a list of users.

.DESCRIPTION
This script takes a list of user emails and automatically provisions OneDrive personal sites for each user. It handles errors gracefully and confirms the creation of OneDrive sites.

.NOTES
Author: John Bignold
Requires: This script requires the SharePoint Online Management Shell module. To install, use `Install-Module -Name Microsoft.Online.SharePoint.PowerShell`
#>


# List of users for whom we want to provision OneDrive personal sites
$userEmails = @(
    'username@dognville.com.au',
    'username2@ognville.com.au'
)

foreach ($user in $userEmails) {
    try {
        # Set the error action preference to 'Stop' to ensure errors are caught
        $ErrorActionPreference = "Stop"

        # Provision OneDrive personal site for the current user
        Request-SPOPersonalSite -UserEmails @($user)  # Note the array notation

        Write-Host "OneDrive personal site provision request submitted successfully for $user!" -ForegroundColor Green

    } catch {
        # Handle any errors for the current user
        Write-Host "Error encountered during provisioning $($_.Exception.Message)" -ForegroundColor Red

    } finally {
        # Reset error action preference
        $ErrorActionPreference = "Continue"
    }
}

# Verify that OneDrive has been created for users
Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"