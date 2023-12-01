<#
.SYNOPSIS
Intalls Office 365 services for PowerShell client.

.NOTES
Author: John Bignold

.LINK
- Connect-MsolService: https://docs.microsoft.com/powershell/module/msonline/connect-msolservice
- Connect-AzureAD: https://docs.microsoft.com/powershell/module/azuread/connect-azuread
- Connect-ExchangeOnline: https://docs.microsoft.com/powershell/module/exchange/connect-exchangeonline
- Connect-MicrosoftTeams: https://docs.microsoft.com/powershell/module/exchange/connect-microsoftteams
#>


function InstallService {
    # Define function parameters. If AudibleAlert is not specified, it defaults to true.
    param (
        [switch]$AudibleAlert = $true
    )

    # Check if the script is being run with administrator privileges.
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host -f Red "Please run the script as an Administrator."
        return
    }

    # Check the current execution policy and prompt to change if necessary.
    $currentPolicy = Get-ExecutionPolicy
    if ($currentPolicy -ne 'RemoteSigned') {
        $userChoice = Read-Host -Prompt "The script requires the execution policy to be set to 'RemoteSigned'. Do you want to proceed? (Y/N)"
        if ($userChoice -eq 'Y') {
            Set-ExecutionPolicy RemoteSigned -Confirm:$false
        } else {
            Write-Host -f Red "Script execution stopped by user."
            return
        }
    }

    # Nested function to check and install PowerShell modules if they are missing.
    function Install-ModuleIfMissing {
        param (
            [string]$ModuleName
        )

        # Check if the module is available on the system.
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            try {
                # Attempt to install the module.
                Install-Module -Name $ModuleName -Scope CurrentUser -Force
                Write-Host -f Green "Installed module: $ModuleName"
            } catch {
                # Handle any installation errors.
                Write-Host -f Red "Failed to install $ModuleName. Error: $_"
            }
        } else {
            Write-Host -f Yellow "Module $ModuleName is already installed" 
        }
    }

    # Notify the user about the installation process.
    Write-Host -BackgroundColor DarkYellow "Installing and importing Azure Management Modules, please wait..." -ForegroundColor Black
    
    # Audible notification using Text-to-Speech if AudibleAlert is set to true.
    if ($AudibleAlert) {
        Function Speak-Text($Text) {
            # Load necessary assembly for speech synthesis.
            Add-Type -AssemblyName System.speech
            $TTS = New-Object System.Speech.Synthesis.SpeechSynthesizer
            $TTS.Speak($Text)
            $TTS.Dispose() # Close the TTS object after using it.
        }
        Speak-Text "Installing and importing Azure Management Modules, please wait..."
    }
    
    Write-Host "`n"
    
    # Install the specified modules.
    Install-ModuleIfMissing -ModuleName 'AzureAD'
    Install-ModuleIfMissing -ModuleName 'ExchangeOnlineManagement'
    Install-ModuleIfMissing -ModuleName 'MSOnline'
    Install-ModuleIfMissing -ModuleName 'MicrosoftTeams'

    # Inform the user about successful module installation.
    Write-Host -f Green "`nModules have been installed successfully."
    Write-Host -f Yellow "Testing module connectivity, please wait...`n"

    # Define the list of modules to test.
    $ModuleNames = @('AzureAD', 'ExchangeOnlineManagement', 'MSOnline', 'MicrosoftTeams')

    # Iterate over each module and attempt to load it, reporting success or failure.
    foreach ($module in $ModuleNames) {
        try {
            Import-Module $module -ErrorAction Stop
            Write-Host -f Green "$module loaded successfully."
        } catch {
            Write-Host -f Red "Failed to load $module. Error: $_"
        }
    }

    Write-Host -f Yellow "`nUpdating PowerShell to Single-Threaded Apartment (STA) mode`n"
    
    # Open a YouTube link (can be modified or removed as per requirement).
    start-process https://www.youtube.com/watch?v=MOXVp8rQYZc

    # Enabling PowerShell Single-Threaded Apartment (STA) mode
    powerShell -STA
}

# Function to uninstall modules (TESTING PURPOSES)
Function RemoveModules {
    # Modules to be removed
    $modulesToRemove = @('ExchangeOnlineManagement', 'MSOnline', 'MicrosoftTeams', 'AzureAD')

    # Iteriate through each to uninstall
    foreach ($module in $modulesToRemove) {
        $modulePath = (Get-Module -ListAvailable | Where-Object Name -eq $module).Path
        if ($modulePath) {
            Remove-Item -Path $modulePath -Recurse -Force
            Write-Host -f Green "$module was removed manually."
        } else {
            Write-Host -f Yellow "$module directory not found."
        }
    }
}


