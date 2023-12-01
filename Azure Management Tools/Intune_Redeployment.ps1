<#
.SYNOPSIS
This script sets the execution policy for the current user, collects a specified user's UPN, 
retrieves the user's Object ID from Azure AD, and removes the user's app data from the Intune 
Management Extension registry.

.EXAMPLE
To run the script, execute it in PowerShell with the necessary administrative privileges.

.NOTES
Author: John Bignold
- The user running the script must have permissions to modify the registry and access Azure AD user information.
- The script assumes that the Intune Management Extension and Azure AD PowerShell Module are installed and configured correctly.

#>



Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Collect User UPN
$userUPN = Read-Host "Enter the UPN of the user"

# Convert UPN to User Object ID
$user = Get-MsolUser -UserPrincipalName $userUPN
$userObjectId = $user.ObjectId

# Delete all apps for a user
$Path = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
Get-Item  -Path $Path\$UserObjectID | Remove-Item -Recurse -Force