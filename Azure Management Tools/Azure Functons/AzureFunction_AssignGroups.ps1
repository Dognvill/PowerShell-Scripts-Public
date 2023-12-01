<#
.SYNOPSIS
This script assigns users to Azure AD and Exchange Online distribution groups.

.DESCRIPTION
AssignGroups is a PowerShell function that automates the process of adding a user to specified Azure AD and/or Exchange Online distribution groups. It checks for a connection to Azure AD and Exchange Online, prompts for the user principal name, and iterates through group assignments based on user input.

.PARAMETER none
The function does not accept parameters and instead relies on interactive input.

.EXAMPLE
PS> AssignGroups

This example runs the AssignGroups function, which will prompt for input and add the specified user to the groups provided.

.NOTES
Author: John Bignold

Make sure you are connected to Azure AD and Exchange Online before running this script.
The script assumes that you have the necessary permissions to modify group memberships in both Azure AD and Exchange Online.
The script will fetch all available groups from Azure AD and Exchange Online to present the most updated list to the administrator.

.LINK
Documentation related to the used cmdlets can be found at:
- Get-AzureADUser: https://docs.microsoft.com/powershell/module/azuread/get-azureaduser
- Get-DistributionGroup: https://docs.microsoft.com/powershell/module/exchange/get-distributiongroup
- Add-DistributionGroupMember: https://docs.microsoft.com/powershell/module/exchange/add-distributiongroupmember
- Add-AzureADGroupMember: https://docs.microsoft.com/powershell/module/azuread/add-azureadgroupmember
#>

Function AssignGroups {
    # Check Azure AD connection
    try {
        Get-AzureADTenantDetail | Out-Null
        Write-Host "Connected to Azure AD." -ForegroundColor Green
    }
    catch {
        Write-Host "Not connected to Azure AD. Please connect and try again." -ForegroundColor Red
        return
    }

    # Check Exchange Online connection
    try {
        Get-EXOMailbox -ResultSize 1 | Out-Null
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Host "Not connected to Exchange Online. Please connect and try again." -ForegroundColor Red
        return
    }

    Write-Host "Welcome to group allocation, follow prompts below." -ForegroundColor Cyan

    # Get user details
    try {
        $Target = Read-Host "Please enter user principal name (e.g., test.user@domain.com)"
        $TargetObj = Get-AzureADUser -ObjectID $Target
    }
    catch {
        Write-Host "User does not exist or an error occurred. Please verify the user principal name and try again." -ForegroundColor Red
        return
    }

    # Fetch all groups from both services before entering the loop
    $allExchangeGroups = Get-DistributionGroup -ResultSize Unlimited
    $allAzureADGroups = Get-AzureADGroup -All $true

    # Loop to add groups
    do {
        $groupName = Read-Host "Enter the name of the group or type 'STOP' to end selection"

        if ($groupName -eq "STOP") {
            Write-Host "Closing group selection..." -ForegroundColor Yellow
            break
        }

        # Filter from the pre-fetched list
        $exchangeGroupObj = $allExchangeGroups | Where-Object { $_.Name -eq $groupName }
        $azureGroupObj = $allAzureADGroups | Where-Object { $_.DisplayName -eq $groupName }

        if ($exchangeGroupObj) {
            Write-Host "[!] - Adding $($TargetObj.UserPrincipalName) to Exchange group $($exchangeGroupObj.Name) ... " -ForegroundColor Yellow -NoNewline
            try {
                Add-DistributionGroupMember -Identity $exchangeGroupObj.Identity -Member $TargetObj.UserPrincipalName
                Write-Host "Done (Exchange)" -ForegroundColor Green
            }
            catch {
                Write-Host "Error adding user to Exchange group: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        elseif ($azureGroupObj) {
            Write-Host "[!] - Adding $($TargetObj.UserPrincipalName) to Azure AD group $($azureGroupObj.DisplayName) ... " -ForegroundColor Yellow -NoNewline
            try {
                Add-AzureADGroupMember -ObjectId $azureGroupObj.ObjectId -RefObjectId $TargetObj.ObjectId
                Write-Host "Done (Azure AD)" -ForegroundColor Green
            }
            catch {
                Write-Host "Error adding user to Azure AD group: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Group not found in both Exchange Online and Azure AD. Please try again." -ForegroundColor Red
        }

        # Ask if user wants to add more groups
        $addMoreGroups = Read-Host "Do you want to add more groups? (y/n)"
    } while ($addMoreGroups -eq "y")

    # Confirm allocation
    Write-Host "Group allocation completed`n" -ForegroundColor Green
    Write-Host "Please review groups below:" -ForegroundColor Yellow
    Get-AzureADUserMembership -ObjectId $TargetObj.ObjectId | Format-Table -Property DisplayName
}

