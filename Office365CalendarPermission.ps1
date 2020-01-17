param (
    [string]   $CredentialName = "Office365",

    [string]   $EmailFrom      = "Calendar Script <CalendarScript@OrganizationName.com>",
    [string]   $EmailTo        = "You <You@OrganizationName.com>",
    [string]   $SMTPServer     = "OrganizationName-com.mail.protection.outlook.com" # Or, if your Exchange is still on-prem/you have a SMTP relay on-prem, change to the appropriate server
)

$ErrorActionPreference = "Stop"

try
{
    Import-Module CredentialManager

    $Head = "<style>table, td, th { margin: auto; border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; } th { padding: 7px; background-color:lightblue; } td { padding: 5px; text-align: center; }</style>"

    $EmptyTable = [pscustomobject]@{
            Identity = "n/a"
            User = "n/a"
            AccessRights = "n/a"
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

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking

    $AllMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox
    $AllCalendars = $AllMailboxes | foreach { $_.Alias + ":\Calendar" }
    $AllCalendarPermissions = $AllCalendars | Get-MailboxFolderPermission -User Default 
    
    # Get all calendar permissions for Default with more than one entry for manual review
    $ReviewPermissions = $AllCalendarPermissions | ? {$_.AccessRights.Count -ne 1}

    if ($ReviewPermissions.Count -eq 0)
    {
        $ReviewPermissions = $EmptyTable
    }

    $ReviewPermissionsTable = $ReviewPermissions | Select-Object Identity, User, AccessRights

    # Get all calendar permissions for Default with a single entry where that entry is not
    $UpdatePermissions = $AllCalendarPermissions | ? {$_.AccessRights.Count -eq 1 -and $_.AccessRights -notcontains "Reviewer"}
    $UpdatePermissions | Set-MailboxFolderPermission -User Default -AccessRights Reviewer

    if ($UpdatePermissions.Count -eq 0)
    {
        $UpdatePermissions = $EmptyTable
    }

    $UpdatePermissionsTable = $UpdatePermissions | Select-Object Identity, User, AccessRights

    $ReviewPermissionsTableHTML = $ReviewPermissionsTable | ConvertTo-Html -PreContent "<center><h2>Review Permissions</h2></center>" -Fragment
    $Body = $UpdatePermissionsTable | ConvertTo-Html -Head $Head -Body $ReviewPermissionsTableHTML -PreContent "<center><h2>Updated Permissions</h2></center>" | Out-String

    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "O365 Calendar permission report" -Body $Body -BodyAsHtml -SmtpServer $SMTPServer

    Remove-PSSession $Session
}
catch
{
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "O365 auto-license failed" -Body $PSItem -SmtpServer $SMTPServer
}