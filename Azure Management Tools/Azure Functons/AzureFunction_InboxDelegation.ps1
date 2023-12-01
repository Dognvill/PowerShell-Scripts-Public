<#
.SYNOPSIS
    Grants specified access permissions (Full Access, Send As, Send On Behalf, or All Access) for a delegate on a given mailbox.

.DESCRIPTION
    This script prompts the user for a mailbox address, delegate address, and the level of access required. It validates the 
    access level input and then assigns the appropriate permissions to the delegate for the mailbox specified.

.PARAMETER mailbox
    The email address of the mailbox that you want to grant access to.

.PARAMETER delegate
    The email address of the user who will be granted access to the mailbox.

.PARAMETER accessLevel
    The level of access to grant: 'FullAccess', 'SendAs', 'SendOnBehalf', or 'AllAccess'.

.NOTES
Author: John Bignold
Prerequisites:  Exchange Online PowerShell module or Exchange Management Shell with appropriate permissions.
#>



# Prompt the user for mailbox and delegate details
$mailbox = Read-Host "Enter the email address of the mailbox you want to grant access to"
$delegate = Read-Host "Enter the email address of the user you want to grant access"
$accessLevel = Read-Host "Enter the access level ('FullAccess', 'SendAs', 'SendOnBehalf', or 'AllAccess')"

# Validate the access level
if ($accessLevel -notin @('FullAccess', 'SendAs', 'SendOnBehalf', 'AllAccess')) {
    Write-Host "Invalid access level specified. Please choose 'FullAccess', 'SendAs', 'SendOnBehalf', or 'AllAccess'."
    exit
}

# Delegate access based on the access level
switch ($accessLevel) {
    'FullAccess' {
        Add-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All -Confirm:$false
        Write-Host "Successfully granted Full Access permission for $delegate on $mailbox"
    }
    'SendAs' {
        Add-RecipientPermission -Identity $mailbox -Trustee $delegate -AccessRights SendAs -Confirm:$false
        Write-Host "Successfully granted Send As permission for $delegate on $mailbox"
    }
    'SendOnBehalf' {
        Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo $delegate
        Write-Host "Successfully granted Send on Behalf permission for $delegate on $mailbox"
    }
    'AllAccess' {
        Add-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess, SendAs, SendOnBehalf -InheritanceType All -Confirm:$false
        Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo $delegate
        Write-Host "Successfully granted all access permission for $delegate on $mailbox"
    }
}

