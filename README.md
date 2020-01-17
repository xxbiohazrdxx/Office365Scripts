# Office365Scripts
Several useful O365 Powershell scripts for organizations.

## Office365CalendarPermissions
The "Default" permission for calendars controls what permissions other users within your oganization have.

This script changes "Default" to grant reviewer rights. This allows the entire organization to view calendar details for all entries not marked as private.

The script also builds a list of users with "Default" permissions that have additional complexity so that they can be manually reviewed and cleaned up if needed.

### Required Modules

This script depends on the CredentialManager module to store and retrieve credentials. It can be installed with the following command in Powershell:

```powershell
Install-Module CredentialManager
```

### Usage

By default, this script will attempt to load a credential named "Office365" from the Credential Manager. However, you may specify a different credential name using the `CredentialName` parameter.

In all cases, if the specified credential name is not found in the Credential Manager, you will be prompted for credentials which will then be saved.

```powershell
.\Office365CalendarPermission.ps1 -CredentialName "Office365" -EmailFrom "LicenseScript@OrganizationName.com" -EmailTo "Admin@OrganizationName.com" -SMTPServer "mail.OrganizationName.com"
```

## Office365AutoLicense

This script is designed to automatically license new users for Office 365 and provision their OneDrive.

### Required Modules

In addition to the CredentialManager module mentioned above, this script also depends on two modules provided by Microsoft for connecting to AzureAD and SharePoint Online. These can be installed with the following commands in PowerShell:

```powershell
Install-Module CredentialManager
Install-Module AzureAD
Install-Module Microsoft.Online.SharePoint.PowerShell
``` 
  
### Usage

As above, credentials will be automatically loaded from the Credential Manager if possible.

If you need to locate your Account SKU you can issue the following commands in Powershell:

```powershell
Connect-MsolService
Get-MsolAccountSku | ft AccountSkuId, SkuPartNumber
```

For example, E3 is `ENTERPRISEPACK` and the complete SKU ID would be `OrganizationName:ENTERPRISEPACK`

To get a list of available plans within your SKU:

```powershell
(Get-MsolAccountSku | Where { $_.SkuPartNumber -eq "ENTERPRISEPACK" }).ServiceStatus
```

You can then provide these plans as an array of strings using the `DisabledPlans` parameter if you do not want them to be active for your users.

```powershell
.\Office365AutoLicense.ps1 -AccountSkuID "OrganizationName:SkuID" -DisabledPlans = @() -SPOServiceURL "https://OrganizationName-admin.sharepoint.com" -EmailFrom "LicenseScript@OrganizationName.com" -EmailTo "Admin@OrganizationName.com" -SMTPServer "mail.OrganizationName.com"
```

Report emails are sent without authentication. Meaning you either need to have an anonymous on-prem relay configured (if you are on a hybrid Exchange environment) or [Direct Send configured with your Exchange Online instance](https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3).
