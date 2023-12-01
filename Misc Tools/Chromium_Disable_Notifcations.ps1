<#
.SYNOPSIS
Configures registry settings to disable web notifications in Google Chrome and Microsoft Edge.

.DESCRIPTION
This script checks for the existence of registry keys that control web notifications settings in Google Chrome and Microsoft Edge under the current user profile. If the keys do not exist, they are created. It then sets the "DefaultNotificationsSetting" property to '2' in both browsers to disable notifications for all websites.

.NOTES
Author: John Bignold
#>


# Checks if the registry key for Chrome and Edge policies exist 
# If the key does not exist, the script creates it with the -Force and -ea SilentlyContinue options.
if ((Test-Path -LiteralPath "Registry::\HKEY_CURRENT_USER\Software\Policies\Google\Chrome") -ne $true) { 
    New-Item "Registry::\HKEY_CURRENT_USER\Software\Policies\Google\Chrome" -force -ea SilentlyContinue 
}

if ((Test-Path -LiteralPath "Registry::\HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge") -ne $true) { 
    New-Item "Registry::\HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge" -force -ea SilentlyContinue 
}

# Sets the value of the "DefaultNotificationsSetting" key to 2 for Chrome to disable notifications for all websites.
New-ItemProperty -LiteralPath 'Registry::\HKEY_CURRENT_USER\Software\Policies\Google\Chrome' -Name 'DefaultNotificationsSetting' -Value '2' -PropertyType DWord

# Sets the value of the "DefaultNotificationsSetting" key to 2 for Edge to disable notifications for all websites.
New-ItemProperty -LiteralPath 'Registry::\HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge' -Name 'DefaultNotificationsSetting' -Value '2' -PropertyType DWord

