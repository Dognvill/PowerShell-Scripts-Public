<#
.SYNOPSIS
This PowerShell script automates the process of downloading files from a specified SharePoint Online document library, uploading them to Azure Blob Storage, and then deleting the local copies. It also logs all operations, including downloads, uploads, deletions, and errors, into a CSV file.

.DESCRIPTION
The script connects to a SharePoint Online site and Azure account, retrieves all files from the specified SharePoint document library, and downloads them to a local directory. Each file is then uploaded to a designated container in Azure Blob Storage. After a successful upload, the local copy of the file is deleted. The script keeps a log of all these operations, noting each file's download, upload, and deletion status, along with any errors encountered, in a CSV file. The log includes the date and time of the operation, the event type, success or failure status, file path, and an internal message describing the event. The script also provides visual feedback in the PowerShell console, using colored messages to distinguish between different types of events.

.PARAMETERS
$SiteURL: URL of the SharePoint site.
$ListName: Name of the SharePoint document library.
$storageAccountName: Azure Storage account name.
$containerName: Name of the Azure Blob Storage container.
$sasToken: Shared Access Signature (SAS) token for Azure Blob Storage access.
$DownloadPath: Local path for downloading files from SharePoint.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
This script does not generate any pipeline output. However, it logs operational details in a CSV file and provides console output.

.EXAMPLE
To execute the script, set the parameters with your SharePoint site URL, list name, Azure Storage account details, and local download path. Run the script in PowerShell with required permissions to access SharePoint and Azure resources.

.NOTES
Created by: John Bignold
Requires: Az Modules, PnP PowerShell, and proper permissions to sources.
#>


# Parameters
$SiteURL = "SiteURL"
$ListName = "Shared Documents"
$storageAccountName = "storageAccountName"
$containerName = "containerName"
$sasToken = "sasToken"
$DownloadPath = "C:\Export"
$LogPath = "C:\AzureUploadLog.txt"

# Connect to externals, SharePoint + Azure
Connect-PnPOnline $SiteURL -Interactive
Connect-AzAccount
 
# User notification
Write-host -f Yellow "Preparing data migration for SharePoint site: $siteURL"
Write-Host "`n"

# Get PnP SharePoint data points (Web + List)
$Web = Get-PnPWeb
$List = Get-PnPList -Identity $ListName
 
# Get all Items from the Library - with progress bar
$global:counter = 0
$ListItems = Get-PnPListItem -List $ListName -PageSize 500 -Fields ID -ScriptBlock { Param($items) $global:counter += $items.Count; Write-Progress -PercentComplete `
    ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting Items from List:" -Status "Processing Items $global:Counter to $($List.ItemCount)"; } 
Write-Progress -Activity "Completed Retrieving Folders from List $ListName" -Completed
 
# Get all Subfolders of the library
$SubFolders = $ListItems | Where { $_.FileSystemObjectType -eq "Folder" -and $_.FieldValues.FileLeafRef -ne "Forms" }
$SubFolders | ForEach-Object {
    try {
        # Ensure All Folders in the Local Path
        $LocalFolder = $DownloadPath + ($_.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)) -replace "/", "\"
        # Create Local Folder, if it doesn't exist
        If (!(Test-Path -Path $LocalFolder)) {
            New-Item -ItemType Directory -Path $LocalFolder | Out-Null
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error while getting subfolders: $($_.Exception.Message)"
        Exit
    }
}

# Get all Files from the folder
$FilesColl = $ListItems | Where { $_.FileSystemObjectType -eq "File" }

# Create Azure Storage Context
$context = New-AzStorageContext -StorageAccountName $storageAccountName -SasToken $sasToken

# User notification
Write-host -f Yellow "Starting data migration for SharePoint site: $siteURL"
Write-Host "`n"
 
# Check if the log file exists, if not create it and add the headers
if (-not (Test-Path -Path $LogPath)) {
    "Date,Event,Success/Fail,FilePath,Internal Message" | Out-File -FilePath $LogPath
}

# Function to properly format CSV data
function Format-CSVField {
    param([string]$field)
    return '"' + $field.Replace('"', '""') + '"'
}

# Function to log events
function Log-Event {
    param(
        [string]$event,
        [string]$status,
        [string]$filePath,
        [string]$internalMessage
    )
    $csvLine = "$(Format-CSVField (Get-Date)),$event,$status,$(Format-CSVField $filePath),$(Format-CSVField $internalMessage)"
    Add-Content $LogPath $csvLine

    # Determine the color based on the event type
    switch ($event) {
        "Download" { $color = "Green" }
        "Upload" { $color = "Cyan" }
        "Delete" { $color = "Magenta" }
        "NotFound" { $color = "Red" }
        "Error" { $color = "Red" }
        Default { $color = "White" }
    }

    # Output the message with the specified color
    Write-Host $internalMessage -ForegroundColor $color
}

# Iterate through each file
$FilesColl | ForEach-Object {
    try {
        # Define variables for SharePoint and Azure paths
        $SharePointRelativePath = $_.FieldValues.FileRef.Substring($Web.ServerRelativeUrl.Length)
        $FileDownloadDir = Join-Path -Path $DownloadPath -ChildPath ($SharePointRelativePath -replace "/", "\").Replace($_.FieldValues.FileLeafRef, '')
        $FileDownloadPath = Join-Path -Path $FileDownloadDir -ChildPath $_.FieldValues.FileLeafRef

        # Download file from SharePoint
        Get-PnPFile -ServerRelativeUrl $_.FieldValues.FileRef -Path $FileDownloadDir -FileName $_.FieldValues.FileLeafRef -AsFile -force
        $downloadMessage = "Downloaded File: $FileDownloadPath"
        Log-Event -event "Download" -status "Success" -filePath $FileDownloadPath -internalmessage $downloadMessage

        # Check if the file exists
        if (-not (Test-Path -Path $FileDownloadPath)) {
            $notFoundMessage = "File not found: $FileDownloadPath"
            Log-Event -event "NotFound" -filePath $FileDownloadPath -internalmessage $notFoundMessage
            continue
        }

        # Define blob name with folder structure
        $blobName = $SharePointRelativePath -replace "^/", "" -replace "\\", "/" -replace "^Shared Documents/", ""

        # Upload file to Azure Blob Storage
        $BlobUpload = Set-AzStorageBlobContent -File $FileDownloadPath -Container $containerName -Blob $blobName -Context $context -Force
        $uploadMessage = "Uploaded File: $blobName"
        Log-Event -event "Upload" -status "Success" -filePath $FileDownloadPath -internalmessage $uploadMessage

        # Delete local temporary file
        Remove-Item -Path $FileDownloadPath -Force
        $removeMessage = "File removed from local host: $blobName"
        Log-Event -event "Delete" -status "Success" -filePath $FileDownloadPath -internalmessage $removeMessage
        Write-Host "`n"

    } catch {
        $ErrorMessage = $_.Exception.Message
        Log-Event -event "Error" -status "Fail" -filePath $FileDownloadPath -internalmessage "Error while processing file: $ErrorMessage"
    }
}

# Cleanup and exit
Disconnect-PnPOnline
Disconnect-AzAccount
