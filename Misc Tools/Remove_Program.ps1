<#
.SYNOPSIS
    Uninstalls a specified program from a Windows computer.

.DESCRIPTION
    This PowerShell function searches for a program's uninstallation information in the system registry and uninstalls it. 
    It allows the user to interactively select a program if multiple matches are found. The function can terminate related 
    processes to ensure a clean uninstallation.

.EXAMPLE
    Uninstall-Program -Interactive
    Runs the uninstallation process in interactive mode, prompting the user for each step.

.NOTES
Author: John Bignold
#>



function Uninstall-Program {
    # Binding cmdlet parameters and setting a confirm impact level
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [switch]$Interactive = $false
    )

    # Start of the uninstallation process
    Write-Host -f Green "Howdy, partner! Welcome to Sheriff John's Uninstallation Saloon!`n"
    
    # Max number of attempts for user to enter a valid program name
    $maxAttempts = 5
    
    # Define the registry paths where uninstallation info is typically stored
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    # Loop through the number of attempts
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $programName = Read-Host "Tell me, what's the varmint of a program you're lookin' to run outta town?"
        Write-Host -f Yellow "Roundin' up the usual suspects and lookin' for '$programName'."

        # Search for the program's uninstallation keys in the registry
        $uninstallKeys = $registryPaths |
        Foreach-Object { Get-ChildItem $_ } |
        Where-Object { $_.GetValue("DisplayName") -like "*$programName*" }

        # If found, exit loop
        if ($uninstallKeys.Count -gt 0) { 
            $DisplayName = $uninstallKeys[0].GetValue("DisplayName")
            Write-Host -f Green "Gotcha! We've cornered '$DisplayName'.`n"

            # Confirm with user
            $validResponses = @("y", "Y", "yes", "YES", "Yes")
            $confirmation = Read-Host "Is this the culprit? Shall we proceed with the eviction? (y/n)"
            if ($confirmation -notin $validResponses) {
                Write-Output "Well alright then, we'll let this one ride off into the sunset."
                return
            }

            break 
        }
        else {
            # Notify user if program not found
            Write-Warning "Can't find any program by the name '$programName' in these parts."
            if ($attempt -lt $maxAttempts) {
                Write-Output "Let's not get discouraged now. Round $attempt of $maxAttempts."
            }
            else {
                Write-Error "Seems we've hit a dead end, partner."
                return
            }
        }
    }

    # If multiple matches are found, ask user to specify
    if ($uninstallKeys.Count -gt 1) {
        $uninstallKeys | foreach { Write-Output "[$_]: $($_.GetValue('DisplayName'))" }

        do {
            $index = Read-Host "Seems we've got a few outlaws. Which one we goin' after?"
        } until ($index -ge 0 -and $index -lt $uninstallKeys.Count)

        # Determine the correct uninstall string based on the user's selection
        $uninstallKey = if ($uninstallKeys[$index].GetValue("QuietUninstallString")) {
            $uninstallKeys[$index].GetValue("QuietUninstallString")
        }
        else {
            $uninstallKeys[$index].GetValue("UninstallString")
        }

    }
    else {
        # Determine the correct uninstall string when only one match is found
        $uninstallKey = if ($uninstallKeys[0].GetValue("QuietUninstallString")) {
            $uninstallKeys[0].GetValue("QuietUninstallString")
        }
        else {
            $uninstallKeys[0].GetValue("UninstallString")
        }
    }

    
    Write-Host -f Yellow "`nChecking the outlaw's hideout..."

    # Extract InstallLocation for the selected program from the uninstall registry key
    if ($uninstallKeys.Count -gt 1) {
        $installLocation = $uninstallKeys[$index].GetValue("InstallLocation")
    }
    else {
        $installLocation = $uninstallKeys[0].GetValue("InstallLocation")
    }

    # Array to accumulate processes that may need to be terminated
    $processesToKill = @()

    # Detect processes based on MainWindowTitle or process name
    $originalDetectedProcesses = Get-Process | Where-Object { $_.MainWindowTitle -like "*$programName*" }
    $processesToKill += $originalDetectedProcesses

    # If InstallLocation is found, determine associated processes
    if ($installLocation) {
        $processesFromInstallLocation = Get-Process | Where-Object { $_.Path -like "$installLocation*" }
        $processesToKill += $processesFromInstallLocation
    }

    # Remove any duplicate processes
    $uniqueProcessesToKill = $processesToKill | Select-Object -Unique

    # List and terminate associated processes before uninstallation
    if ($uniqueProcessesToKill) {
        Write-Host -f Yellow "Found these no-goodniks related to $DisplayName"
        $uniqueProcessesToKill | ForEach-Object { Write-Output "$($_.Id) - $($_.Name) - $($_.Path)" }
        Write-Host -f Yellow "`nRounding up the culprits...`n"
        $uniqueProcessesToKill | Stop-Process -Force
    }

    # Execute the uninstall command
    if ($PSCmdlet.ShouldProcess($programName, "Uninstall")) {
        if ($uninstallKey) {
            if ($Interactive) {
                $args = "/c $uninstallKey"
            }
            else {
                if ($uninstallKeys[0].GetValue("QuietUninstallString")) {
                    $args = "/c $uninstallKey"
                }
                else {
                    $args = "/c $uninstallKey /s /VERYSILENT"
                }
            }
            Write-Host -f Yellow "Hold on to your hats, we're kickin' out $DisplayName..."
            Start-Process -FilePath "cmd.exe" -ArgumentList $args -Wait
            SLEEP 3
        }
        else {
            Write-Error "Looks like $DisplayName gave us the slip, partner. Can't find the info we need."
        }
    
        Write-Host -f Green "Alright, we've successfully run $DisplayName out of town!"
    }
}
