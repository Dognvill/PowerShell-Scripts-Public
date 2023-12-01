<#
.SYNOPSIS
Connects to an Azure AD tenant and exports user details to an Excel file.

.DESCRIPTION
This script connects to Azure AD using a specific Tenant ID, retrieves the user principal names and display names of all users, exports the data to an Excel file named after the tenant, and then disconnects from Azure AD.

.EXAMPLE
PS> .\ExportAadUsers.ps1
This will execute the script for each tenant ID specified in the $tenantIds array.

.NOTES
Author: John
Requires: The AzureAD PowerShell module.

#>

# Array of tenant IDs you want to connect to
$tenantIds = @('TENANT_ID')

foreach ($tenantId in $tenantIds) {
    # Connect to Azure AD with specific Tenant ID
    Connect-AzureAD -TenantId $tenantId
    
    # Retrieve tenant name
    $tenantName = (Get-AzureADTenantDetail).DisplayName

    # Construct file name
    $fileName = "C:\${tenantName}Export.xlsx"

    # Fetch UPNs and DisplayNames
    $users = Get-AzureADUser | Select-Object UserPrincipalName, DisplayName

    # Export to Excel
    $users | Export-Excel -Path $fileName -WorksheetName "UserDetails"
    
    # Disconnect from the current Azure AD session
    Disconnect-AzureAD
}
