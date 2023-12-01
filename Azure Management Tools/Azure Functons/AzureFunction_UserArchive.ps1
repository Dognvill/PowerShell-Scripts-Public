<#
.SYNOPSIS
    This function archives a specified user by exporting their data from the Microsoft Compliance Admin Centre.

.DESCRIPTION
    ArchiveUser connects to the Microsoft Compliance Admin Centre, initiates a compliance search for the given user, 
    and attempts to export the search results. 

    It handles errors and retries if the export is not immediately available, and opens relevant URLs in web browsers
    for manual follow-up if the automated process fails after a number of attempts.

.PARAMETER user
    The email address of the user to archive. It is obtained interactively during function execution.

.NOTES
Author: John Bignold
Prerequisites:  Microsoft Compliance Admin Centre access, appropriate permissions, and necessary modules (IPPSSession)

#>

Function ArchiveUser {
    Write-Host -BackgroundColor DarkYellow "`n `n Connecting to Microsoft Compliance Admin Centre, please login when prompted... `n `n " -ForegroundColor Black
    Function Speak-Text($Text) { Add-Type -AssemblyName System.speech; $TTS = New-Object System.Speech.Synthesis.SpeechSynthesizer; $TTS.Speak($Text) }
    Speak-Text "Connecting to Microsoft Compliance Admin Centre, please login when prompted"

    # Connect to Compliance
    Connect-IPPSSession

    Write-Host -BackgroundColor DarkYellow "`n Welcome to John's impeccable and delectable archival script, please follow the instructions below to begin the process of archiving a user!" -ForegroundColor Black
    Function Speak-Text($Text) { Add-Type -AssemblyName System.speech; $TTS = New-Object System.Speech.Synthesis.SpeechSynthesizer; $TTS.Speak($Text) }
    Speak-Text "Welcome to John's impeccable and delectable archival script, please wait whilst we export the user!"

    # Collect user address and start compliance search
    $user = Read-Host
    $export_user = $user + "_Export"
    New-ComplianceSearch $user -ExchangeLocation $user | Start-ComplianceSearch
    Write-Host "Starting compliance search for $user" -ForegroundColor Yellow


    # Checking Compliance search is ready and loops until it's avaliable
    $maxRetries = 5
    $retryCount = 0
    $success = $false

    while (($retryCount -lt $maxRetries) -and (!$success)) {
        try {
            New-ComplianceSearchAction $user -Export -Format Fxstream -ErrorAction Stop
            $success = $true
        }
        catch {
            Write-Host "Waiting for export to be avaliable... Please refer to Sherksophone for further assistance." -ForegroundColor Yellow
            Start-Process "https://www.youtube.com/watch?v=pxw-5qfJ1dk&t=78s"
            Start-sleep 60
            $retryCount++
        }
    }

    if (!$success) {
        Write-Host "The export failed after $maxRetries attempts. Please try again later." -ForegroundColor Red
    }
    else {
        Write-Host "The export was successful." -ForegroundColor Green
        Write-Host "Collecting URL for export request." -ForegroundColor Yellow
    }

    # Collect export URl with error checking like above
    while (($retryCount -lt $maxRetries) -and (!$success)) {
        try {
            Get-ComplianceSearchAction $export_user -IncludeCredential | FL -ErrorAction Stop
            $success = $true
        }
        catch {
            Write-Host "Waiting for export to be avaliable... Please refer to Lil Nas X, whilst we search your export." -ForegroundColor Yellow
            Start-Process https://www.youtube.com/watch?v=r7qovpFAGrQ
            Start-sleep 120
            $retryCount++
        }
    }

    if (!$success) {
        Write-Host "The export failed after $maxRetries attempts. Please export manually..."
    }
    else {
        Write-Host  -BackgroundColor DarkYellow "The export was successful, please visit Compliance Admin Centre to download the exported .pst"  -ForegroundColor Black
        Write-Host  -BackgroundColor DarkYellow "Remember to remove from Cove Backup"  -ForegroundColor Black
        Write-Host  -BackgroundColor DarkYellow "Openning..."  -ForegroundColor Black
        Start-Process MicrosoftEdge -ArgumentList "https://compliance.microsoft.com/contentsearchv2?viewid=export"
    }

    Start-Sleep -Seconds 5

}

