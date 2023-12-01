# Welcome intro
$t = @"

███╗░░░███╗░█████╗░██╗██╗░░░░░  ███████╗██╗░░██╗░█████╗░██╗░░██╗░█████╗░███╗░░██╗░██████╗░███████╗
████╗░████║██╔══██╗██║██║░░░░░  ██╔════╝╚██╗██╔╝██╔══██╗██║░░██║██╔══██╗████╗░██║██╔════╝░██╔════╝
██╔████╔██║███████║██║██║░░░░░  █████╗░░░╚███╔╝░██║░░╚═╝███████║███████║██╔██╗██║██║░░██╗░█████╗░░
██║╚██╔╝██║██╔══██║██║██║░░░░░  ██╔══╝░░░██╔██╗░██║░░██╗██╔══██║██╔══██║██║╚████║██║░░╚██╗██╔══╝░░
██║░╚═╝░██║██║░░██║██║███████╗  ███████╗██╔╝╚██╗╚█████╔╝██║░░██║██║░░██║██║░╚███║╚██████╔╝███████╗
╚═╝░░░░░╚═╝╚═╝░░╚═╝╚═╝╚══════╝  ╚══════╝╚═╝░░╚═╝░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░╚══╝░╚═════╝░╚══════╝
"@

for ($i = 0; $i -lt $t.length; $i++) {
    if ($i % 2) {
        $c = "red"
    }
    elseif ($i % 5) {
        $c = "yellow"
    }
    elseif ($i % 7) {
        $c = "yellow"
    }
    else {
        $c = "Yellow"
    }
    write-host $t[$i] -NoNewline -ForegroundColor $c
}
Write-Host "`n"
Write-host -BackgroundColor DarkYellow "Welcome to Mail Exchange, please wait whilst we retrieve mailbox permissions, forwarding rules, and distribution group members" -ForegroundColor Black


# Get all user and shared mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
Write-Host -f Yellow "Processing mailbox permissions...`n"

# Get mailbox permissions
$mailboxPermissions = @()
foreach ($mailbox in $mailboxes) {
    Write-Host "Processing mailbox permissions for:" -NoNewline
    Write-Host -f Cyan " $($mailbox.PrimarySmtpAddress)"
    $permissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName | Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false }
    $mailboxPermissions += $permissions | Select-Object @{Name = "Mailbox"; Expression = { $mailbox.PrimarySmtpAddress } }, User, AccessRights
}

# Export mailbox permissions to CSV
if ($PSScriptRoot -ne "" -and (Test-Path $PSScriptRoot)) {
    $mailboxPermissions | Export-Csv -Path "$PSScriptRoot\MailboxPermissions.csv" -NoTypeInformation
    Write-Host -f Green "Successfully exported mailbox permissions to $PSScriptRoot\MailboxPermissions.csv"
}
else {
    $mailboxPermissions | Export-Csv -Path "C:\MailboxPermissions.csv" -NoTypeInformation
    Write-Warning "Unable to export to $PSScriptRoot directory. Exporting to C:\MailboxPermissions.csv instead."
    Write-host -f Green "Successfully exported"
}
Write-Host "================================================`n"

# Get forwarding rules
Write-Host -f Yellow "Processing forwarding rules...`n"
Start-Sleep -Seconds 3
$forwardingRules = @()
foreach ($mailbox in $mailboxes) {
    Write-Host "Processing forwarding rules for:" -NoNewline
    Write-Host -f cyan " $($mailbox.PrimarySmtpAddress)"
    if ($mailbox.ForwardingSmtpAddress -ne $null) {
        $forwardingRules += New-Object -TypeName PSObject -Property @{
            Mailbox                    = $mailbox.PrimarySmtpAddress
            ForwardingSmtpAddress      = $mailbox.ForwardingSmtpAddress
            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        }
    }
}

# Export forwarding rules to CSV
if ($PSScriptRoot -ne "" -and (Test-Path $PSScriptRoot)) {
    $forwardingRules | Export-Csv -Path "$PSScriptRoot\ForwardingRules.csv" -NoTypeInformation
    Write-Host -f Green "Successfully exported forwarding rules to $PSScriptRoot\ForwardingRules.csv"
}
else {
    $forwardingRules | Export-Csv -Path "C:\ForwardingRules.csv" -NoTypeInformation
    Write-Warning "Unable to export to $PSScriptRoot directory. Exporting to C:\ForwardingRules.csv instead."
    Write-host -f Green "Successfully exported"
}
Write-Host "================================================`n"

# Get distribution groups and members
Write-Host -f Yellow "Processing distribution groups and memberships...`n"
Start-Sleep -Seconds 3
$distGroups = Get-DistributionGroup -ResultSize Unlimited
$distGroupMembers = @()
foreach ($group in $distGroups) {
    Write-Host "Processing distribution group members for:" -NoNewline
    Write-Host -f cyan " $($group.Name)"
    $members = Get-DistributionGroupMember -Identity $group.DistinguishedName
    foreach ($member in $members) {
        $distGroupMembers += New-Object -TypeName PSObject -Property @{
            GroupName = $group.Name
            Member    = $member.PrimarySmtpAddress
        }
    }
}

# Export distribution group members to CSV
if ($PSScriptRoot -ne "" -and (Test-Path $PSScriptRoot)) {
    $distGroupMembers | Export-Csv -Path "$PSScriptRoot\DistributionGroupMembers.csv" -NoTypeInformation
    Write-Host -f Green "Successfully exported distribution group members to $PSScriptRoot\DistributionGroupMembers.csv"
}
else {
    $distGroupMembers | Export-Csv -Path "C:\DistributionGroupMembers.csv" -NoTypeInformation
    Write-Warning "Unable to export to $PSScriptRoot directory. Exporting to C:\DistributionGroupMembers.csv instead."
    Write-host -f Green "Successfully exported"
}
Write-Host "================================================`n"



