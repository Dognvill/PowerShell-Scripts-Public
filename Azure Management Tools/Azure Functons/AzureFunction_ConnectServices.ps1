<#
.SYNOPSIS
Connects to Office 365 services for specified tenants.

.DESCRIPTION
The ConnectService function provides an interactive prompt to select a predefined tenant or manually 
enter the tenant's Global Admin credentials. It then establishes a connection to MS Online Service, 
Azure AD, and Exchange Online for the chosen tenant.

.NOTES
Author: John Bignold

Ensure that the MSOnline, AzureAD, and ExchangeOnline PowerShell modules are installed and imported into your session before running this script.
The script assumes that you have the necessary permissions to connect to the services as a Global Admin for the tenants.

.LINK
- Connect-MsolService: https://docs.microsoft.com/powershell/module/msonline/connect-msolservice
- Connect-AzureAD: https://docs.microsoft.com/powershell/module/azuread/connect-azuread
- Connect-ExchangeOnline: https://docs.microsoft.com/powershell/module/exchange/connect-exchangeonline
#>

function ConnectService {

    #Function Mission-Impossible { Musical tones }
    do {
        # Tenant selection
        Write-Host "`n============= Select Tenant Client =============="    -ForegroundColor Yellow
        Write-Host "`ta. '1' for Tenant 1'"                                 -ForegroundColor Yellow
        Write-Host "`tb. '2' for Tenant 2'"                                 -ForegroundColor Yellow
        Write-Host "`tb. '3' for Tenant 3'"                                 -ForegroundColor Yellow
        Write-Host "`tb. '4' for Tenant 4'"                                 -ForegroundColor Yellow
        Write-Host "`tb. '5' to enter tenant manually"                      -ForegroundColor Yellow
        Write-Host "==============================================="        -ForegroundColor Yellow
        $Role = Read-Host "`Please confirm selection"

    } until (($Role -eq '1') -or ($Role -eq '2') -or ($Role -eq '3') -or ($Role -eq '4') -or ($Role -eq '5'))
    switch ($Role) {
        '1' {
            $acctName = "admin@dognville.com.au"
        }
        '2' {
            $acctName = "admin@dognville.com.au"
        }
        '3' {
            $acctName = "admin@dognville.com.au"
        }
        '4' {
            $acctName = "admin@dognville.com.au"
        }
        '5' {
            $acctName = Read-Host "Enter tenant Global Admin"
        }
    }

    # Connect Services
    Connect-MsolService
    Connect-AzureAD
    Connect-ExchangeOnline -UserPrincipalName $acctName -ShowProgress $true

    # Welcome user

        $t = @"

░█████╗░░█████╗░███╗░░██╗███╗░░██╗███████╗░█████╗░████████╗███████╗██████╗░
██╔══██╗██╔══██╗████╗░██║████╗░██║██╔════╝██╔══██╗╚══██╔══╝██╔════╝██╔══██╗
██║░░╚═╝██║░░██║██╔██╗██║██╔██╗██║█████╗░░██║░░╚═╝░░░██║░░░█████╗░░██║░░██║
██║░░██╗██║░░██║██║╚████║██║╚████║██╔══╝░░██║░░██╗░░░██║░░░██╔══╝░░██║░░██║
╚█████╔╝╚█████╔╝██║░╚███║██║░╚███║███████╗╚█████╔╝░░░██║░░░███████╗██████╔╝
░╚════╝░░╚════╝░╚═╝░░╚══╝╚═╝░░╚══╝╚══════╝░╚════╝░░░░╚═╝░░░╚══════╝╚═════╝░
"@

    for ($i = 0; $i -lt $t.length; $i++) {
        if ($i % 2) {
            $c = "DarkCyan"
        }
        elseif ($i % 5) {
            $c = "green"
        }
        elseif ($i % 7) {
            $c = "green"
        }
        else {
            $c = "green"
        }
        write-host $t[$i] -NoNewline -ForegroundColor $c
    }

    Function Speak-Text($Text) { Add-Type -AssemblyName System.speech; $TTS = New-Object System.Speech.Synthesis.SpeechSynthesizer; $TTS.Speak($Text) }
    Speak-Text "You're now connected to all Azure Services, please use the command line to add or remove users"
    Write-Host -BackgroundColor DarkYellow "You're now connected to all Azure Services, please use the command line to add or remove users (AddUser - RemoveUser)" -ForegroundColor Black
    # Mission-Impossible
}
