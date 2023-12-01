<#
.SYNOPSIS
    Uploads a file to an FTP server using the WinSCP .NET assembly.

.DESCRIPTION
    This script sets up a connection to an FTP server using explicit TLS security and uploads a specified file.
    It requires the WinSCP .NET assembly to be installed and accessible on the local system.

.PARAMETER HostName
    The hostname or IP address of the FTP server.

.PARAMETER UserName
    The username for authentication with the FTP server.

.PARAMETER Password
    The password for authentication with the FTP server.

.PARAMETER sourcePath
    The local path of the file to upload.

.PARAMETER remotePath
    The remote FTP directory path where the file will be uploaded.

.NOTES
    Ensure that WinSCPnet.dll is present at the specified path and that the credentials and paths are correct.
    Using GiveUpSecurityAndAcceptAnyTlsHostCertificate makes the script accept any TLS/SSL certificate,
    which could pose a security risk in a production environment.

.LINK
    For more information on WinSCP .NET assembly, visit: https://winscp.net/eng/docs/library
#>

# Load WinSCP .NET assembly
# Please ensure the path points to the WinSCPnet.dll in your local WinSCP installation
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll" # Specify path here

# Set FTP details
$HostName = "FTP.Server.Address" #

# Load WinSCP .NET assembly
# Please ensure the path points to the WinSCPnet.dll in your local WinSCP installation
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll" # Specify path here

# Set FTP details
$HostName = "FTP.Sever.Address" # FTP server address
$UserName = "username"             # Specify your FTP username
$Password = "password"             # Specify your FTP password

# Set up session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol  = [WinSCP.Protocol]::Ftp
    HostName  = $HostName
    UserName  = $UserName
    Password  = $Password
    FtpSecure = [WinSCP.FtpSecure]::Explicit
}

# Instruct WinSCP to accept any certificate
$sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true

$session = New-Object WinSCP.Session

try {
    # Connect to FTP server
    $session.Open($sessionOptions)

    # Upload files
    $sourcePath = "C:\testfile.txt" # Specify the path to your source file
    $remotePath = "/uploadfilepath/" # Specify the path to your remote folder

    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

    $transferResult = $session.PutFiles($sourcePath, $remotePath, $False, $transferOptions)

    # If there are any errors, an exception is thrown
    $transferResult.Check()

    # Print results
    foreach ($transfer in $transferResult.Transfers) {
        Write-Host ("Upload of {0} succeeded" -f $transfer.FileName)
    }

}
catch {
    # Handle exceptions
    Write-Host ("Error: {0}" -f $_.Exception.Message)
    exit 1
}
finally {
    # Disconnect and clean up
    $session.Dispose()
}
