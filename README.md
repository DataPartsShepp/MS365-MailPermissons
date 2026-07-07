# Mail Permissions Report

This repository contains a PowerShell script to connect to Microsoft 365 via app-only authentication, read a list of mailbox addresses, and export permission details to an Excel workbook.

## Files

- `Get-MailboxPermissionReport.ps1` - main script
- `connection.json` - sample app-only connection configuration
- `connection.example.json` - generic example configuration
- `connection.certificate.example.json` - certificate-based configuration example
- `emails.txt` - sample mailbox list

## Prerequisites

- PowerShell 7.x or later is recommended, but Windows PowerShell 5.1 works if `ExchangeOnlineManagement` and `ImportExcel` modules are supported.
- An Azure AD app registration with Exchange Online app-only permission and certificate authentication:
  - certificate thumbprint (`CertificateThumbprint`) when `UseCertificate` is set to `true`
  - client secret (`AppSecret`) may work only if your installed `ExchangeOnlineManagement` module supports it, but certificate auth is recommended
- `ExchangeOnlineManagement` PowerShell module
- `ImportExcel` PowerShell module

The script installs missing modules automatically if they are not already present.

> Note: `ExchangeOnlineManagement` 3.x and later no longer supports app secret authentication via `AppSecret`/`ClientSecret` for app-only Exchange Online connections. Certificate auth is the supported and recommended option.

### Required app permissions

The app registration must be granted app-only permissions for Exchange Online and should be consented by an administrator.

- `Exchange.ManageAsApp` for the **Office 365 Exchange Online** API (not Microsoft Graph)

This permission is found under:
- Azure AD > App registrations > your app > API permissions > Add a permission
- Select `Office 365 Exchange Online` (or `Exchange`), then `Application permissions`
- Choose `Exchange.ManageAsApp`

Admin consent must be granted for the permission so the script can authenticate app-only and query mailbox/group data without interactive sign-in.

> Note: This script uses Exchange Online cmdlets such as `Connect-ExchangeOnline`, `Get-Mailbox`, `Get-UnifiedGroup`, and `Get-UnifiedGroupLinks`, so Graph permissions are not required for the current implementation.

## Configuration

Create or update `connection.json` with your tenant and app credentials. For certificate auth, use the dedicated example file `connection.certificate.example.json`.

```json
{
  "TenantId": "your-tenant-id.onmicrosoft.com",
  "AppId": "your-app-registration-client-id",
  "AppSecret": null,
  "UseCertificate": true,
  "CertificateThumbprint": "your-certificate-thumbprint"
}
```

This example uses certificate auth, which is the recommended method for current ExchangeOnlineManagement module versions.

### Finding your certificate thumbprint

Use the Windows certificate manager or PowerShell to locate the cert thumbprint:

- Open `certmgr.msc` and check the `Personal\Certificates` store.
- Select the certificate, open `Details`, then find the `Thumbprint` field.

Or run in PowerShell:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint
```

Copy the printed thumbprint value into `connection.json`.

## Input list

Use a plain text file or CSV file to provide mailboxes or Microsoft 365 group SMTP addresses.

### Plain text example

`emails.txt`

```text
user1@domain.com
user2@domain.com
group1@domain.com
```

### CSV example

`emails.csv`

```csv
Email
user1@domain.com
group1@domain.com
```

## Run the script

```powershell
.\Get-MailboxPermissionReport.ps1 \
  -ConfigPath .\connection.json \
  -EmailListPath .\emails.txt \
  -OutputPath .\MailboxPermissionsReport.xlsx
```

## Output

The script writes permission data to an Excel workbook with the following worksheets:

- `MailboxPermissions` - mailbox-level access entries
- `FolderPermissions` - folder-level access entries for common folders
- `GroupPermissions` - Microsoft 365 group membership and owner details

## Notes

- Mailbox entries that cannot be found are skipped with a warning.
- The script currently checks `Inbox`, `Calendar`, `Contacts`, `Tasks`, and `Sent Items` for folder permissions.
- Group entries are identified using `Get-UnifiedGroup` and exported under `GroupPermissions`.

## Troubleshooting

- If `ExchangeOnlineManagement` fails to connect, verify the app registration permissions and tenant values.
- If `ImportExcel` is missing, the script installs it automatically.
- If the CSV file does not contain an `Email` column, the script will fail with a validation message.
