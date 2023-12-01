
<#
.SYNOPSIS
    Adds a list of users to a specified email group.

.DESCRIPTION
    This script takes an array of email addresses and adds them to a specified distribution group
    in Microsoft Exchange. It uses the Add-DistributionGroupMember cmdlet to add each user.

.PARAMETER GroupEmail
    The email address of the group to which members will be added. This should be replaced
    with the actual group email address before running the script.

.PARAMETER EmailAddresses
    An array of user email addresses that will be added to the email group. Update this list
    with the actual email addresses you wish to add to the group.

.EXAMPLE
    .\AddMembersToGroup.ps1
    This example adds the predefined list of users in the $emailAddresses array to the group
    defined in $groupEmail.

.NOTES
Author: John Bignold

.LINK
    For more information about Add-DistributionGroupMember, visit:
    https://docs.microsoft.com/powershell/module/exchange/add-distributiongroupmember

#>

# Replace GroupEmailAddress with the email address of the email group you want to add members to.
$groupEmail = "GroupEmailAddress"

# Define an array of email addresses to add to the email group
$emailAddresses = @("username@domain.com.au", "username1@domain.com.au")

# Iterate through each email address and add it to the email group
foreach ($email in $emailAddresses) {
    try {
        # Add the email to the group using the Add-DistributionGroupMember cmdlet
        Add-DistributionGroupMember -Identity $groupEmail -Member $email
        Write-Host "Added $email to the group $groupEmail"
    }
    catch {
        Write-Host "Failed to add $email to the group $groupEmail. Error message: $($_.Exception.Message)"
    }
}

