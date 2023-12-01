
<#
.SYNOPSIS
This script removes all licenses from specified users in Office 365.

.DESCRIPTION
This PowerShell script connects to Microsoft Online Service and iteratively removes all licenses
from each user specified by the administrator running the script. The script uses display names
to find the corresponding UserPrincipalName for each user and then removes the licenses.

.NOTES
Author: John Bignold
#>

# Install Module
Import-Module MSOL

# Connect to Service
Connect-MsolService

# Number of users to be processed
$userCount = Read-Host "Enter number of users to be processed"

# Loop through the number of users
for ($i = 1; $i -le $userCount; $i++) {

    # Collect ser display name
    $userDisplayName = Read-Host "Enter Display Name of the user $i"

    # Converts display name to UPN
    $userUPN = (Get-MsolUser -SearchString $userDisplayName).UserPrincipalName

    # Remove licenses of the user
    (get-MsolUser -UserPrincipalName $userUPN).licenses.AccountSkuId |
    foreach {
        Set-MsolUserLicense -UserPrincipalName $userUPN -RemoveLicenses $_
    }
    Write-Host "Successfully removed all licenses from user: $userUPN"
}