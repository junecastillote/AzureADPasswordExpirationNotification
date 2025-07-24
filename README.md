# AzureADPasswordExpirationNotification

![PowerShell](https://img.shields.io/badge/PowerShell-Module-blue.svg)
![Platform](https://img.shields.io/badge/platform-AzureAD-blue)
![License](https://img.shields.io/badge/license-MIT-green.svg)

**AzureADPasswordExpirationNotification** (a.k.a. *AzAdPwdExpAlert*) is a PowerShell module to automate proactive password expiration notices for users in Azure AD (via Microsoft Graph). It retrieves upcoming password expirations, notifies users ahead of time, and optionally sends summary reports to admins.

---

## Features ✉️

* **Get-UserPasswordExpiration**: Fetches user password expiry dates and days remaining.
* **Send-UserPasswordExpirationNotice**: Sends customized HTML email notices to users and summary reports to admins.
* Multiple notifications per user (e.g., at 15, 10, 7, 5, 3, and 1 days before expiry).
* Optionally redirect notification delivery (useful for testing).
* Support for email attachments and copying HTML notifications to disk.
* Admin summary report via HTML template.

---

## Prerequisites

* PowerShell 5.1+ / PowerShell Core 7+
* Microsoft.Graph PowerShell SDK:

  ```powershell
  Install-Module Microsoft.Graph
  ```

* Azure AD App Registration with application permissions for automated / unattended jobs:

  * `Mail.Send`
  * `Organization.Read.All`
  * `User.Read.All`

---

## Installation

1. [Download the zip file](https://github.com/junecastillote/AzureADPasswordExpirationNotification/archive/refs/heads/main.zip) or clone the repository.
2. Extract (if ZIP download) to your preferred folder.
3. Import the module in PowerShell.

```powershell
Import-Module .\AzureADPasswordExpirationNotification\AzAdPwdExpAlert.psd1
```

---

## Usage

### 1. Connect to Microsoft Graph

```powershell
# Choose a command below depending on whether you're using delegated or application log-in.

# Delegated user sign-in
Connect-MgGraph -Scopes User.Read.All, Organization.Read.All, Mail.Send -TenantId "tenant.onmicrosoft.com"

# Application sign-in with application ID and certificate
Connect-MgGraph -TenantId "tenant.onmicrosoft.com" -ClientId "client id here" -CertificateThumbPrint "certificate thumbprint here"

# Application sign-in with secret credential
$secretCredential = Get-Credential # Input the client id as username and secret key as the password.
Connect-MgGraph -TenantId "tenant.onmicrosoft.com" -SecretCredential $secretCredential
```

### 2. Check password expirations

```powershell
$expirations = Get-UserPasswordExpiration
```

### 3. Notify users and/or admins

```powershell
$expirations |
  Send-UserPasswordExpirationNotice `
    -NotifyUsers `
    -From "noreply@contoso.com" `
    -SendReportToAdmins "admin@contoso.com"
```

### Parameters

| Parameter                           | Type     | Required | Description                                              |
| ----------------------------------- | -------- | -------- | -------------------------------------------------------- |
| `-InputObject`                      | Object   | Yes      | The output object result of `Get-UserPasswordExpiration` |
| `-PasswordNotificationWindowInDays` | Int      | No       | Days before expiry to notify (default: 15,10,7,5,3,1)    |
| `-RedirectNotificationTo`           | String[] | No       | Redirect all emails to test mailbox                      |
| `-Attachment`                       | String[] | No       | Full path of files to attach (if any)                    |
| `-CopyHtmlToFolder`                 | String   | No       | Save HTML versions of emails to disk                     |
| `-EmailTemplate`                    | String   | No       | Path to custom HTML user notification template           |
| `-From`                             | String   | No       | Sender email address                                     |
| `-NotifyUsers`                      | Switch   | No       | Send the email notification to users if used             |
| `-SendReportToAdmins`               | String[] | No       | A collection of administrator recipient email addresses  |
| `-Subject`                          | String   | No       | Custom email subject                                     |

---

## Examples

**Basic Notification**

```powershell
Get-UserPasswordExpiration |
  Send-UserPasswordExpirationNotice `
    -NotifyUsers `
    -From "noreply@contoso.com"
```

**Custom Notification Window with Audit Copy**

```powershell
Get-UserPasswordExpiration |
  Send-UserPasswordExpirationNotice `
    -NotifyUsers `
    -From "noreply@contoso.com" `
    -PasswordNotificationWindowInDays 7,3 `
    -CopyHtmlToFolder "C:\\Temp\\Notices" `
    -SendReportToAdmins "helpdesk@contoso.com"
```

**Test Mode (Redirect Emails)**

```powershell
Get-UserPasswordExpiration |
  Send-UserPasswordExpirationNotice `
    -NotifyUsers `
    -From "noreply@contoso.com" `
    -RedirectNotificationTo "tester@contoso.com"
```

---

## HTML Templates

Templates are found in the `resource/` directory:

* `user_notification_template.html`: Template for individual user notifications
* `admin_report_template.html`: Template for summary report

Variables:

* **User template**: `$name`, `$expireInDays`, `$expirationDate`, `$upn`
* **Admin report**: `$organizationName`, `$totalCount`, `$notifiedCount`, `$notNotifiedCount`

---

## Module Manifest

```powershell
@{
  ModuleVersion = '0.0.2'
  Author = 'June Castillote'
  FunctionsToExport = @(
    'Get-UserPasswordExpiration',
    'Send-UserPasswordExpirationNotice'
  )
  RequiredModules = @('Microsoft.Graph')
}
```

---

## License

Licensed under the [MIT License](LICENSE).

---

## Author

**June Castillote**
[GitHub](https://github.com/junecastillote)
[june.castillote@gmail.com](mailto:june.castillote@gmail.com)

---
