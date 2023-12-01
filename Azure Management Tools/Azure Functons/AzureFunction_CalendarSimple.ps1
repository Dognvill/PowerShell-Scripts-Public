# List of users to grant permissions
$usersToGrant = @(
    'username1@dognvill.com.au',
    'username2@dognvill.com.au'
)

# List of mailboxes to grant access permissions too
$mailboxesToGrant = @(
    'username1@dognvill.com.au',
    'username2@dognvill.com.au'
)
# Ask the user for the type of access rights
$accessRights = Read-Host "Please enter the access rights (e.g., Owner, Reviewer, Contributor, etc.)"

# Validate the input for AccessRights
if (-not [System.Enum]::IsDefined([Microsoft.Exchange.WebServices.Data.FolderPermissionLevel], $accessRights)) {
    Write-Host -ForegroundColor Red "Invalid Access Rights. Please enter a valid permission level."
    return
}

# Grant specified permission to specified mailboxes for each user
foreach ($user in $usersToGrant) {
    Write-Host -ForegroundColor Cyan "Processing user: $user"
    
    # Loop through each mailbox in the mailboxesToGrant array
    foreach ($mailboxEmail in $mailboxesToGrant) {
    
        if ($mailboxEmail -ne $user) {
            # The calendar folder path
            $calendarPath = $mailboxEmail + ":\Calendar"
            Write-Host -ForegroundColor Yellow "    Checking permissions for: $calendarPath"
            
            # Check if the user already has permissions
            try {
                $permissionsGet = Get-MailboxFolderPermission -Identity $calendarPath -User $user -ErrorAction Stop
                Write-Host -ForegroundColor Yellow "    User already has permissions. Updating to $accessRights..."
                # If this succeeds, the user has permissions and we can set them
                $PermissionsSet = Set-MailboxFolderPermission -Identity $calendarPath -User $user -AccessRights $accessRights
                Write-Host -ForegroundColor Green "    Successfully updated permissions to $accessRights."
            } catch {
                Write-Host -ForegroundColor Red "    No existing permission entry found for user. Adding permissions..."
                # If the command fails, it means the user does not have permissions and we need to add them
                $PermissionsAdd = Add-MailboxFolderPermission -Identity $calendarPath -User $user -AccessRights $accessRights
                Write-Host -ForegroundColor Green "    Successfully added permissions to $accessRights."
            }
            
            # Confirm permissions lines
            Write-Host -ForegroundColor Yellow "    Confirming granted permissions..."
            # Execute the command to get the permissions and print them
            Get-MailboxFolderPermission -Identity $calendarPath -User $user
            Write-Host "`n"
            
        } else {
            Write-Host "Skipping permission grant to self."
        }
    }
}