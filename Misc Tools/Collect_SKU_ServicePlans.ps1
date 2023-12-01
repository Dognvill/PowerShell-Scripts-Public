<#
.SYNOPSIS
Retrieves all subscribed SKUs from Azure AD and lists out each SKU's service plans.

.EXAMPLE
To run the script, execute it in a PowerShell session that has the AzureAD module installed. You must be signed into an account with permissions to view SKUs and service plans.

.NOTES
Author: John Bignold
- Requires the AzureAD PowerShell module.
- The user must have administrative privileges to access SKUs information in Azure AD.
- Output is stored in the `$LICARRAY` array, which can be used for further processing or output.
#>



# Connect to Azure Active Directory
Connect-AzureAD

# Retrieve all subscribed SKUs in the Azure AD tenant
$ALLSKUS = Get-AzureADSubscribedSku

# Initialize an array to store license information
$LICARRAY = @()

# Iterate through each SKU and collect service plan details
foreach ($SKU in $ALLSKUS) {
    $servicePlans = $SKU.ServicePlans | ForEach-Object {
        "Service Plan Name: $($_.ServicePlanName), Provisioning Status: $($_.ProvisioningStatus)"
    }
    
    # Add the SKU part number and service plans to the array
    $LICARRAY += "SERVICE PLAN: " + $SKU.SkuPartNumber
    $LICARRAY += $servicePlans
    $LICARRAY += "" # Add an empty string for a separator between SKUs
}

# Output the license array
$LICARRAY
