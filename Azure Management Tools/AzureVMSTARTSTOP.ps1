<#
.SYNOPSIS
Manages the power state of a specified Azure VM using Managed Service Identities.

.DESCRIPTION
This script handles the starting or stopping of an Azure VM based on its current power state. 
It connects to Azure using Managed Service Identities (MSI) - either system-assigned (SA) or 
user-assigned (UA). Upon establishing a connection, it checks the VM's status and performs the 
action required to change its state to either start or stop.

.PARAMETER resourceGroup
Specifies the name of the resource group where the target VM is located.

.PARAMETER VMName
Specifies the name of the VM to be managed.

.PARAMETER method
Determines the method of connection using either "SA" for system-assigned managed identity or "UA" for user-assigned managed identity.

.PARAMETER UAMI
Specifies the name of the user-assigned managed identity, required if the method is "UA".

.EXAMPLE
.\AzureVMSTARTSTOP.ps1 -resourceGroup "ResourceGroup01" -VMName "VM01" -method "UA" -UAMI "MyManagedIdentity"

Connects to Azure using the user-assigned managed identity "MyManagedIdentity" and manages the power state of the VM named "VM01" in the resource group "ResourceGroup01".

.NOTES
Requires: Az Modules, Proper permissions for the managed identity used to execute the script.
#>



Param(
    [string]$resourceGroup,
    [string]$VMName,
    [string]$method,
    [string]$UAMI 
)

$automationAccount = "ACC_NAME"

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process | Out-Null

# Connect using a Managed Service Identity
try {
    $AzureContext = (Connect-AzAccount -Identity).context
}
catch {
    Write-Output "There is no system-assigned user identity. Aborting."; 
    exit
}

# set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription `
    -DefaultProfile $AzureContext

if ($method -eq "SA") {
    Write-Output "Using system-assigned managed identity"
}
elseif ($method -eq "UA") {
    Write-Output "Using user-assigned managed identity"

    # Connects using the Managed Service Identity of the named user-assigned managed identity
    $identity = Get-AzUserAssignedIdentity -ResourceGroupName $resourceGroup `
        -Name $UAMI -DefaultProfile $AzureContext

    # validates assignment only, not perms
    if ((Get-AzAutomationAccount -ResourceGroupName $resourceGroup `
                -Name $automationAccount `
                -DefaultProfile $AzureContext).Identity.UserAssignedIdentities.Values.PrincipalId.Contains($identity.PrincipalId)) {
        $AzureContext = (Connect-AzAccount -Identity -AccountId $identity.ClientId).context

        # set and store context
        $AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext
    }
    else {
        Write-Output "Invalid or unassigned user-assigned managed identity"
        exit
    }
}
else {
    Write-Output "Invalid method. Choose UA or SA."
    exit
}

# Get current state of VM
$status = (Get-AzVM -ResourceGroupName $resourceGroup -Name $VMName `
        -Status -DefaultProfile $AzureContext).Statuses[1].Code

Write-Output "`r`n Beginning VM status: $status `r`n"

# Start or stop VM based on current state
if ($status -eq "Powerstate/deallocated") {
    Start-AzVM -Name $VMName -ResourceGroupName $resourceGroup -DefaultProfile $AzureContext
}
elseif ($status -eq "Powerstate/running") {
    Stop-AzVM -Name $VMName -ResourceGroupName $resourceGroup -DefaultProfile $AzureContext -Force
}

# Get new state of VM
$status = (Get-AzVM -ResourceGroupName $resourceGroup -Name $VMName -Status `
        -DefaultProfile $AzureContext).Statuses[1].Code  

Write-Output "`r`n Ending VM status: $status `r`n `r`n"

Write-Output "Account ID of current context: " $AzureContext.Account.Id