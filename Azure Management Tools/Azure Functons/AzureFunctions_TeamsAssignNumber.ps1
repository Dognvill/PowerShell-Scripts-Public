<#
.SYNOPSIS
The AssignNumber function is designed to streamline the assignment of phone numbers to Microsoft Teams users. 
It automates the process of connecting to the Microsoft Teams admin portal, choosing a dial plan policy, and 
assigning a phone number.

.NOTES
Author: John Bignold
#>

Function AssignNumber {

    # Check if the Microsoft Teams PowerShell module is installed
    if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
        # If the module is not installed, install it
        Install-Module -Name MicrosoftTeams -Confirm:$false
    }

    # Check if the Microsoft Teams PowerShell module is loaded in the current session
    $teamsModule = Get-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue
    if ($teamsModule -eq $null) {
        # If the module is not loaded, load it
        Import-Module MicrosoftTeams
    }


    Write-Host -f Yellow "Checking connection to Teams Admin Portal, please wait..."

    # Check if the user is already connected to Microsoft Teams
    try {
        Get-Team -ErrorAction SilentlyContinue
        Write-Host -f Green "You are connected to Microsoft Teams.`n`n"
    }
    catch {
        # If the user is not connected, run the Connect-MicrosoftTeams cmdlet
        Write-Host -f Yellow "You are not connected to Microsoft Teams. Connecting now..."
        Connect-MicrosoftTeams
    }

    # Display the menu of options
    Write-Host "Select a dial plan policy:"
    Write-Host "  1. AU-NSW-ACT-DP"
    Write-Host "  2. AU-VIC-TAS-DP"
    Write-Host "  3. AU-WA-SA-NT-DP"
    Write-Host "  4. AU-QLD-DP"

    # Prompt the user for a choice
    $validChoices = 1..4
    $choice = Read-Host "Enter a choice (1-4): "
    if ($choice -notmatch "^\d$" -or $validChoices -notcontains $choice) {
        Write-Host "Invalid choice"
        exit
    }

    # Apply the dial plan policy and phone number
    $UserID = Read-Host "Enter UPN"
    $number = Read-Host "Enter whole number (e.g. 61737095589)"

    Write-Host "Processing..." -ForegroundColor Yellow
    Set-CsPhoneNumberAssignment -Identity "$UserID" -EnterpriseVoiceEnabled $true
    Grant-CsTenantDialPlan -Identity "$UserID" -PolicyName "$policyName"
    Set-CsPhoneNumberAssignment -Identity "$UserID" -PhoneNumberType DirectRouting -PhoneNumber "$number"
    Grant-CsOnlineVoiceRoutingPolicy -Identity "$UserID" -PolicyName "Australia"

    # Check if the phone number was applied successfulGet-CsOnlineUser -identity $UserIDly --> USER ONLY
    $phoneNumber = (Get-CsOnlineUser -Identity $UserID).LineUri
    if ($phoneNumber -like "*$number") {
        Write-Host "Phone number applied successfully: $phoneNumber" -ForegroundColor Green
    }
    else {
        Write-Host "Failed to apply phone number" -ForegroundColor Red
    }
}

