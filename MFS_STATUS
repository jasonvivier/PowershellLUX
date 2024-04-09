#Connects to ADDS and gets all enabled users with email addresses
#Connects to Exchange online to removes users with shared mailboxes
#Connect to Azure and gets the authentication method used. If none it assumes MFA enabled is false. 

# Connect to the local Active Directory
Import-Module ActiveDirectory

# Get enabled AD users with an email address
$adUsers = Get-ADUser -Filter 'Enabled -eq $true -and mail -like "*"' -Properties mail

# Connect to Office 365 Exchange
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection
Import-PSSession $ExchangeSession -DisableNameChecking

# Get mailboxes that are not shared and exist in the list of AD users
$nonSharedMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $_.RecipientTypeDetails -ne "SharedMailbox" -and $adUsers.mail -contains $_.PrimarySmtpAddress
}

# Connect to Azure AD for MFA status
Connect-AzureAD
Connect-MsolService

# Initialize the result array
$userDetails = @()

foreach ($mailbox in $nonSharedMailboxes) {
    $userPrincipalName = $mailbox.UserPrincipalName
    $azureUser = Get-AzureADUser -ObjectId $userPrincipalName -ErrorAction SilentlyContinue

    # Skip if the user cannot be found in Azure AD
    if (-not $azureUser) {
        continue
    }

    $msolUser = Get-MsolUser -UserPrincipalName $userPrincipalName
    $authMethods = $msolUser.StrongAuthenticationMethods | ForEach-Object { $_.MethodType }
    $mfaEnabled = $authMethods.Count -gt 0

    # Add user details to the array
    $userDetails += [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        UserPrincipalName = $userPrincipalName
        MFAEnabled = $mfaEnabled
        AuthMethods = ($authMethods -join ', ')
    }
}

# Output the user details
$userDetails | Format-Table -AutoSize

# Clean up the session
Remove-PSSession $ExchangeSession
