
Function AddAdmin{
# Define the username and password for the new admin account
$Username = "Admin"
$password = "Password"

# Create a new local user account
$LocalUser = New-LocalUser -Name $Username -Password (ConvertTo-SecureString $Password -AsPlainText -Force) -Description "Local Administrator Account"

# Add the new local user account to the Administrators group
Add-LocalGroupMember -Group "Administrators" -Member $LocalUser.Name

# Output the result
Write-Host -f Green "New local admin account created:"
Write-Host -f Green "Username: $Username"
Write-Host -f Green "Password: $Password"
}
