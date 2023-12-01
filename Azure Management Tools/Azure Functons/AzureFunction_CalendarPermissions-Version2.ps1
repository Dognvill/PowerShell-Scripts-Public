<#
.SYNOPSIS
This script manages calendar permissions in Exchange Online.

.DESCRIPTION
This PowerShell script provides an interface for adding or removing calendar permissions for users in 
an Exchange Online environment. It prompts the user to connect to Exchange, then allows the user to add 
or remove permissions, and provides feedback on the current permissions set.

.NOTES
Author: John Bignold

#>




$t = @"

░█████╗░░█████╗░██╗░░░░░███████╗███╗░░██╗██████╗░░█████╗░██████╗░
██╔══██╗██╔══██╗██║░░░░░██╔════╝████╗░██║██╔══██╗██╔══██╗██╔══██╗
██║░░╚═╝███████║██║░░░░░█████╗░░██╔██╗██║██║░░██║███████║██████╔╝
██║░░██╗██╔══██║██║░░░░░██╔══╝░░██║╚████║██║░░██║██╔══██║██╔══██╗
╚█████╔╝██║░░██║███████╗███████╗██║░╚███║██████╔╝██║░░██║██║░░██║
░╚════╝░╚═╝░░╚═╝╚══════╝╚══════╝╚═╝░░╚══╝╚═════╝░╚═╝░░╚═╝╚═╝░░╚═╝
"@

for ($i = 0; $i -lt $t.length; $i++) {
    if ($i % 2) {
        $c = "red"
    }
    elseif ($i % 5) {
        $c = "yellow"
    }
    elseif ($i % 7) {
        $c = "yellow"
    }
    else {
        $c = "Yellow"
    }
    write-host $t[$i] -NoNewline -ForegroundColor $c
}
# Connect to Exchange Online
Write-Host "`n `n `n Connecting to Exchange Shell, please login when prompted." -ForegroundColor Yellow
Connect-ExchangeOnline

# While loop for scriptblock
$script = "yes"
While ($script -eq "yes") {
    # Set the initial response to "yes" to start the loop
    Write-Host -BackgroundColor DarkYellow "Do you need to add or remove calendar permissions? (add, remove, exit)" -ForegroundColor Black
    $response = read-host
    # Exit script
    if ($response -eq "exit") {
        $script = "no"
    }

    # Loop while the response is "add"
    while ($response -eq "add") {
        # Prompt the user to enter a user ID (User Principal Name) and add calendar ID variable
        $UserID = read-host "Enter user UPN"
        Write-Host "Processing... `n" -ForegroundColor Yellow
        while (-not(Get-User $UserID)) {
            $UserID = Read-Host "User not found. Enter a valid user ID. `n `n"
        }
        $calendarID = $userID + ":\Calendar"

        # Remove the default permission from the calendar
        Set-MailboxFolderPermission -Identity "$calendarID" -User default -AccessRights None -ErrorAction SilentlyContinue
        Write-Host  -BackgroundColor DarkYellow "Follow prompts below to allocate permissions for 'People in my Organization'." -ForegroundColor Black

        # Allocate permissions for 'People in my organization'
        # Error handling
        $ValidAccessRights = "Reviewer", "Editor", "Author", "PublishingEditor", "Owner", "None"
        $AccessRights = ""

        # Select the access rights for 'People in my organization'
        while ($AccessRights -notin $ValidAccessRights) {
            $AccessRights = Read-Host "Enter the Access Rights (e.g. Reviewer, Editor, None, etc.)"
            if ($AccessRights -notin $ValidAccessRights) {
                Write-Host "Invalid access right entered. Please enter one of the following: `n$($ValidAccessRights -join ', ') `n" -ForegroundColor Red
            }
        }
        
        # Allocate calendar permissions
        Set-MailboxFolderPermission -Identity "$calendarID" -User Default -AccessRights $AccessRights -ErrorAction Inquire
        Write-Host "Permission $AccessRights added for all authenticated users." -ForegroundColor Green

        # Allocate permissions for other authenticated users.
        # Error handling
        Write-Host "`n `n"
        Write-Host -BackgroundColor DarkYellow "Do you need to allocate individual permissions to select user? (y/n)" -ForegroundColor Black
        $answer = Read-Host

        
        while ($answer -eq "y") {
            # Select authenticated user to allocate
            $User = Read-Host "Enter user ID to share $userID calendar with."
            while (-not(Get-User $User)) {
                $User = Read-Host "User not found. Enter a valid user ID to share $userID calendar with."
            }

            $ValidAccessRights = "Reviewer", "Editor", "Author", "PublishingEditor", "Owner", "None"
            $AccessRights = ""
            
            # Select the access rights for authenticated users
            while ($AccessRights -notin $ValidAccessRights) {
                $AccessRights = Read-Host "Enter the Access Rights (e.g. Reviewer, Editor, Author, etc.)."
                if ($AccessRights -notin $ValidAccessRights) {
                    Write-Host "Invalid access right entered. Please enter one of the following: $($ValidAccessRights -join ', ') `n" -ForegroundColor Red
                }
            }

            # Allocate calendar permissions
            Write-Host "Adding requested permissions for authenticated users: $User" -ForegroundColor Yellow
            Write-Host "Processing... `n" -ForegroundColor Yellow
            # Error handling
            try {
                Add-MailboxFolderPermission -Identity "$calendarID" -User $User -AccessRights $AccessRights -ErrorAction Stop
                Write-Host "Permissions added to calendar" -ForegroundColor Green
            }
            catch {
                Write-Host "The specified user already has a permission entry for the folder. Updating permissions..."  -ForegroundColor Red
                Set-MailboxFolderPermission -Identity "$calendarID" -User $User -AccessRights $AccessRights -ErrorAction Ignore
            }

            # Prompt the user to continue or exit the loop
            Write-Host  -BackgroundColor DarkYellow "Enter 'yes' to add another user or 'no' to continue." -ForegroundColor Black
            $script = Read-Host
        }
                        

        # Show the calendar permissions for the user
        Write-Host  -BackgroundColor DarkGreen "Permissions added:"  -ForegroundColor Black
        Get-MailboxFolderPermission -Identity "$CalendarID"

        # Prompt the user to continue or exit the loop
        Write-Host  "`n `n"
        Write-Host  -BackgroundColor DarkYellow "Enter 'yes' to restart script or 'no' to exit." -ForegroundColor Black
        Write-Host  "`n `n"
        $script = Read-Host
        $Response = "no"
    }

    # Loop while the response is "remove"
    while ($response -eq "remove") {
        # Prompt the user to enter a user ID (User Principal Name) and add calendar ID variable
        $UserID = read-host "Enter user UPN"
        Write-Host "Processing... `n" -ForegroundColor Yellow
        while (-not(Get-User $UserID)) {
            $UserID = Read-Host "User not found. Enter a valid user ID. `n `n"
        }
        $calendarID = $userID + ":\Calendar"

        # Show the calendar permissions for the user
        Write-Host -BackgroundColor DarkGreen "Current Permissions added:" -ForegroundColor Black
        Get-MailboxFolderPermission -Identity "$CalendarID"

        # Prompt for user to remove
        $User = Read-Host "Enter user ID to remove from $userID's calendar."
        while (-not(Get-User $User)) {
            $User = Read-Host "User not found. Enter a valid user ID to share $userID calendar with."
        }

        # Remove user from calendar + Error handling
        try {
            Remove-MailboxFolderPermission -Identity "$calendarID" -User $User -ErrorAction Stop
            Write-Host "$user permissions has been successfully removed from $UserID's calendar" -ForegroundColor Green
        }
        catch {
            Write-Host "An error occured, please restart script"  -ForegroundColor Red
        }

        # Show the calendar permissions for the user
        Write-Host  -BackgroundColor DarkGreen "Permissions added:"  -ForegroundColor Black
        Get-MailboxFolderPermission -Identity "$CalendarID"
        
        # Prompt the user to continue or exit the loop
        Write-Host  "`n `n"
        Write-Host  -BackgroundColor DarkYellow "Enter 'yes' to restart script or 'no' to exit." -ForegroundColor Black
        Write-Host  "`n `n"
        $script = Read-Host
        $Response = "no"
    }
}