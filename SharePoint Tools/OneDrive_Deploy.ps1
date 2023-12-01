<#
.SYNOPSIS
This script ensures the installation and operation of the OneDrive client on a Windows device.

.DESCRIPTION
The script performs a series of steps to check the status of OneDrive on the client machine, create necessary directories, download the OneDrive client if not present, install it, and finally, start it using a scheduled task. It also provides a pop-up notification to the end user upon completion.

.NOTES
Author: John Bignold
Compatibility: This script is compatible with Windows environments where the OneDrive client can be installed.
#>


# # STEP 1 # #  
# Checks if OneDrive is already installed and running on the client
function CheckOneDriveStatus {
    
    # Check if OneDrive is installed
    $isInstalled = Test-Path -Path "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    if ($isInstalled) {
        Write-Host -f Green "OneDrive is already installed on the client."
    }
    else {
        Write-Host -f Red "OneDrive is not installed on the client."
    }

    # Check if OneDrive is running
    $isRunning = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($isRunning) {
        Write-Host -f Green "OneDrive is currently running on the client.`n"
    }
    else {
        Write-Host -f Yellow "OneDrive is not currently running on the client.`n"
    }
}

# # STEP 2 # # 
# Checks if Software folder exists
function EnsureDirectoryExists {
    param (
        [string]$path
    )
    if (-not (Test-Path -Path $path)) {
        New-Item -Path $path -ItemType Directory
        Write-Host -f Green "C:\Software successfully created.`n"
    }
    else {
        Write-Host -f Green "C:\Software successfully created.`n"
    }
}

# # STEP 3 # # 
# Downloads OneDrive client to Software folder via 'WebClient Request'
function DownloadOneDrive {
    param (
        [string]$url,
        [string]$destinationPath
    )
    $webClient = New-Object System.Net.WebClient
    try {
        $webClient.DownloadFile($url, $destinationPath)
        Write-Host -f Green "OneDrive Client successfully downloaded.`n"
    }
    catch {
        Write-Host -f Yellow "WebClient download failed. Attempting Invoke-WebRequest instead."
        Invoke-WebRequest -Uri $url -OutFile $destinationPath
    }
}

# # STEP 4 # # 
# Installs OneDrive to client device
function InstallOneDrive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$setupPath,

        [Parameter(Mandatory = $true)]
        [string]$onedriveExe
    )

    # Check if the setup path exists
    if (-not (Test-Path -Path $setupPath)) {
        Write-Host "Failed to find OneDrive setup at path: $setupPath" -ForegroundColor Red
        return
    }

    # Attempt to install OneDrive client silently
    try {
        Write-Host "Installing OneDrive client, please wait..." -ForegroundColor Yellow
        $process = Start-Process -filepath $setupPath -argumentlist "/silent" -NoNewWindow
        start-sleep 10
    }
    catch {
        Write-Host "An error occurred during installation: $_" -ForegroundColor Red
        return
    }

    # Define potential OneDrive executable paths
    $onedrivePaths = @(
        "C:\Program Files\Microsoft OneDrive\OneDrive.exe",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    )

    # Check each path and use the first one that exists
    $onedriveExe = $onedrivePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    # Verify if OneDrive was installed successfully
    if ($onedriveExe) {
        Write-Host "OneDrive has been installed successfully at $onedriveExe." -ForegroundColor Green
    }
    else {
        Write-Host "OneDrive failed to install or was not found in the expected locations." -ForegroundColor Red
    }
}


# # STEP 5 # #
# Runs OneDrive via scheduled task for local user creds
function StartOneDrive {
    # Define potential OneDrive executable paths
    $onedrivePaths = @(
        "C:\Program Files\Microsoft OneDrive\OneDrive.exe",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    )

    # Check each path and use the first one that exists
    $onedriveExe = $onedrivePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($onedriveExe) {
        # Define the task action
        Write-Host -f Yellow "Creating scheduled task action..."
        $action = New-ScheduledTaskAction -Execute $onedriveExe

        # Define the task principal to set the user context and logon type
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive

        # Register the scheduled task
        Write-Host -f Yellow "Running scheduled task action...`n"
        $registerschedule = Register-ScheduledTask -Action $action -Principal $principal -TaskName "RunOneDrive" -Force

        # Start the task to run OneDrive
        $startschedule = Start-ScheduledTask -TaskName "RunOneDrive"
        Write-Host -f Green "OneDrive has been started via scheduled task."
    }
    else {
        Write-Host -f Red "Could not find OneDrive.exe. Ensure it's installed properly."
    }
}


# # STEP 6 # #
# End-user notifcaiton
function Show-PopupNotification {
    param (
        [string]$title,
        [string]$message
    )

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}


### Execution List ###

# # STEP 1 # # 
# Checks if OneDrive client is already installed and running on client
CheckOneDriveStatus

# # STEP 2 # # 
# Checks if Software folder exists
EnsureDirectoryExists -path "C:\software"

# # STEP 3 # # 
# Downloads OneDrive client to Software folder via 'WebClient Request'
DownloadOneDrive -url "https://go.microsoft.com/fwlink/p/?linkid=844652" -destinationPath "C:\software\OneDriveSetup.exe"

# # STEP 4 # # 
# Installs OneDrive to client device
InstallOneDrive -setupPath "C:\software\OneDriveSetup.exe" -onedriveExe "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"

# # STEP 5 # #
#R uns OneDrive via scheduled task for local user creds
StartOneDrive

# # STEP 6 # # 
# OneDrive install notification for end-user
Show-PopupNotification -title "Notification - Microsoft OneDrive" -message "OneDrive installation is complete, please review all staff email sent from admin for further instructions. Have a great day! :)" 



### Other lines for playing with KLM pathways ###
 
$tenantID = "TENANT_ID" #TenantID GUID.
new-item -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "KFMSilentOptIn" -Value $tenantID -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "KFMSilentOptInWithNotification" -Value "1" -Force

$HKLMregistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'##Path to HKLM keys
$DiskSizeregistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\DiskSpaceCheckThresholdMB'##Path to max disk size key
$TenantGUID = 'ffb9ff5c-db27-40f8-925a-e46a9b291a8e'

if (!(Test-Path $HKLMregistryPath)) { New-Item -Path $HKLMregistryPath -Force }
if (!(Test-Path $DiskSizeregistryPath)) { New-Item -Path $DiskSizeregistryPath -Force }

New-ItemProperty -Path $HKLMregistryPath -Name 'SilentAccountConfig' -Value '1' -PropertyType DWORD -Force | Out-Null ##Enable silent account configuration
New-ItemProperty -Path $DiskSizeregistryPath -Name $TenantGUID -Value '102400' -PropertyType DWORD -Force | Out-Null ##Set max OneDrive threshold before prompting

# Define the user's email
$userEmail = "username@dognvill.com.au"

# Create the odopen URL
$odopenURL = "odopen://sync?useremail=$userEmail"

# Start the OneDrive sync process
Start-Process $odopenURL

# Optional: Output to the console that the process has been started
Write-Host "OneDrive sync for $userEmail has been initiated."
