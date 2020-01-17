param (
    [string]   $CredentialName = "Office365",

    [string]   $AccountSkuID   = "OrganizationName:SkuID",
    [string[]] $DisabledPlans  = @(),
    [string]   $SPOServiceURL  = "https://OrganizationName-admin.sharepoint.com",

    [string]   $EmailFrom      = "License Script <LicenseScript@OrganizationName.com>",
    [string]   $EmailTo        = "You <You@OrganizationName.com>",
    [string]   $SMTPServer     = "OrganizationName-com.mail.protection.outlook.com" # Or, if your Exchange is still on-prem/you have a SMTP relay on-prem, change to the appropriate server
)

$ErrorActionPreference = "Stop"

try
{
    Import-Module CredentialManager
    Import-Module AzureAD
    Import-Module Microsoft.Online.SharePoint.PowerShell

    $Head = "<style>table, td, th { margin: auto; border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; } th { padding: 7px; background-color:lightblue; } td { padding: 5px; text-align: center; }</style>"

    $EmptyTable = [pscustomobject]@{
            SignInName = "n/a"
            UserType = "n/a"
            BlockCredential = "n/a"
            CloudExchangeRecipientDisplayType = "n/a"
        }

    $Credentials = Get-StoredCredential -Target $CredentialName

    if ($Credentials -eq $null)
    {
        Write-Host "The credentail store does not have an entry named `"$CredentialName`"."
        Write-Host "Enter your credentials and they will be added to the store:"

        $O365Username = Read-Host "O365 Username: "
        $O365Password = Read-Host "O365 Password: " -AsSecureString
        New-StoredCredential -Target $CredentialName -UserName $O365Username -SecurePassword $O365Password

        $Credentials = Get-StoredCredential -Target $CredentialName
    }

    Connect-MsolService -Credential $Credentials
    Connect-SPOService -Url $SPOServiceURL -Credential $Credentials

    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $DisabledPlans

    # Get unlicensed users who are:
    # 1) UserType is Member                - External users can be added to Teams, etc. and they do not need to be licensed
    # 2) BlockCredential is false          - Termed employees who are blocked in AD do not need to be licensed
    # 3) First/Last names are not null     - Previously used Exchange Recipient Display Type (below), however certain Azure service accounts have a type of '1073741824' and will not be filtered out
    #
    # CloudExchangeRecipientDisplayType - Prevent adding licenses to cloud only accounts, equipment accounts, room accounts, etc.
    # null       - Shared
    # 6          - External
    # 7          - Room Mailbox
    # 8          - Equipment
    # 1073741824 - User
    #
    $UnlicensedUsers = Get-MsolUser -All -UnlicensedUsersOnly -Synchronized | ? {$_.UserType -eq "Member"} | ? {-not $_.BlockCredential} | ? {$_.FirstName -ne $null -and $_.LastName -ne $null}

    # If you're not in the US, change to the appropriate region
    $UnlicensedUsers | Set-MsolUser -UsageLocation US

    # Add licenses
    $UnlicensedUsers | Set-MsolUserLicense -AddLicenses $AccountSkuID -LicenseOptions $LicenseOptions
    
    # Provision OneDrive
    $UnlicensedUsers | % { Request-SPOPersonalSite -UserEmails @($_.UserPrincipalName) }

    if ($UnlicensedUsers.Count -eq 0)
    {
        $UnlicensedUsers = $EmptyTable
    }

    $UnlicensedUsersTable = $UnlicensedUsers | Select-Object SignInName, UserType, BlockCredential, CloudExchangeRecipientDisplayType

    # Get disabled users w/ a standard mailbox
    # These users could be recently disabled and have not had their mailbox exported or converted to shared yet.
    # Add these to the report so the admin is informed of their state and can take manual action to export/delete/convert the mailbox.
    #
    # Note, disabled with a license is a valid configuration in the following instances:
    # 1) Skype/Teams attendant. Will be disabled but have a MCOEV license
    # 2) ????? Probably others, but not discovered yet
    #
    # For this reason, these accounts will always show up in the report
    $ReviewUsers = Get-MsolUser -All | ? {$_.UserType -eq "Member"} | ? {$_.BlockCredential} | ? {$_.IsLicensed} | ? {$_.CloudExchangeRecipientDisplayType -eq 1073741824}

    if ($ReviewUsers.Count -eq 0)
    {
        $ReviewUsers = $EmptyTable
    }

    $ReviewUsersTable = $ReviewUsers | Select-Object SignInName, UserType, BlockCredential, CloudExchangeRecipientDisplayType

    # Get disabled users w/ a converted mailbox
    # These users have been disabled and the admin has converted their mailbox to shared. (CloudExchangeRecipientDisplayType not equal to 1073741824)
    # Because the mailbox has been converted, a license is no longer needed to prevent automatic deletion.
    $DisabledUsers = Get-MsolUser -All | ? {$_.UserType -eq "Member"} | ? {$_.BlockCredential} | ? {$_.IsLicensed} | ? {$_.CloudExchangeRecipientDisplayType -ne 1073741824}
    $DisabledUsers | Set-MsolUserLicense -RemoveLicenses $_.Licenses.AccountSkuId

    if ($DisabledUsers.Count -eq 0)
    {
        $DisabledUsers = $EmptyTable
    }

    $DisabledUsersTable = $DisabledUsers | Select-Object SignInName, UserType, BlockCredential, CloudExchangeRecipientDisplayType

    $UnlicensedUsersTableHTML = $UnlicensedUsersTable | ConvertTo-Html -PreContent "<center><h2>Added Licenses</h2></center>" -Fragment
    $DisabledUsersTableHTML = $DisabledUsersTable | ConvertTo-Html -PreContent "<center><h2>Disabled Licenses</h2></center>" -Fragment
    $Body = $ReviewUsersTable | ConvertTo-Html -Head $Head -Body "$UnlicensedUsersTableHTML $DisabledUsersTableHTML" -PreContent "<center><h2>Review Licenses</h2></center>" | Out-String
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "O365 license report" -Body $Body -BodyAsHtml -SmtpServer $SMTPServer
}
catch
{
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "O365 auto-license failed" -Body $PSItem -SmtpServer $SMTPServer
}