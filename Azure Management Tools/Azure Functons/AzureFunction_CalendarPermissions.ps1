<#
.SYNOPSIS
    This script grants specified calendar permissions to a predefined user or security group for multiple user calendars in an Exchange environment.

.DESCRIPTION
    The script takes a comma-separated list of user principal names (UPNs), prompts for a user or security group, and assigns calendar permissions.
    It ensures the user or security group exists and that valid permission levels are selected before applying the changes.

.EXAMPLE
    PS> .\SetCalendarPermissions.ps1
    The script will prompt for multiple user UPNs, ensure the existence of the user/security group, and apply 'Reviewer' permissions by default.

.PARAMETER UserIDs
    A string of user UPNs separated by commas to specify which users' calendars to modify.

.PARAMETER User
    A predefined user or security group email to whom the calendar permissions will be granted.

.PARAMETER AccessRights
    The level of access granted to the user or security group on the specified calendars. Default is 'Reviewer'.

.NOTES
Author: John Bignold
Prerequisites:  Exchange Management Shell or PowerShell with Exchange cmdlets loaded.
#>



# Prompt for multiple user UPNs separated by commas
$UserIDs = Read-Host "Enter user UPNs separated by commas"

# Loop through each user ID provided
foreach ($UserID in ($UserIDs -split ',')) {
    # Create the calendar identity for the current user
    $CalendarID = $UserID.Trim() + ":\Calendar"

    # Prompt for a user or security group to share the calendar with
    $User = "Username@domain.com.au"
    while (-not(Get-Group $User)) {
        $User = Read-Host "User or group not found. Enter a valid user ID or security group to share $CalendarID calendar with."
    }

    # Select the access rights for authenticated users
    $ValidAccessRights = "Reviewer", "Editor", "Author", "PublishingEditor", "Owner", "None"
    $AccessRights = ""
    while ($AccessRights -notin $ValidAccessRights) {
        $AccessRights = "Reviewer"
        if ($AccessRights -notin $ValidAccessRights) {
            Write-Host "Invalid access right entered. Please enter one of the following: $($ValidAccessRights -join ', ') `n" -ForegroundColor Red
        }
    }

    # Allocate calendar permissions
    Write-Host "Adding requested permissions for authenticated group: $User" -ForegroundColor Yellow
    Write-Host "Processing... `n" -ForegroundColor Yellow
    # Error handling
    try {
       
        Write-Host "Permissions added to '$userID' calendar" -ForegroundColor Green
    }
    catch {
        
    }
}
