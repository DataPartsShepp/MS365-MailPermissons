# Mail Permissions Report

This repository contains a PowerShell script to connect to Microsoft 365 via app-only authentication, read a list of mailbox addresses, and export permission details to an Excel workbook.

## Files

- `Get-MailboxPermissionReport.ps1` - main script
- `connection.json` - sample app-only connection configuration
- `emails.txt` - sample mailbox list

## Prerequisites

- PowerShell 7.x or later is recommended, but Windows PowerShell 5.1 works if `ExchangeOnlineManagement` and `ImportExcel` modules are supported.
- An Azure AD app registration with Exchange Online app-only permission and either:
  - client secret (`AppSecret`), or
  - certificate thumbprint (`CertificateThumbprint`) when `UseCertificate` is set to `true`
- `ExchangeOnlineManagement` PowerShell module
- `ImportExcel` PowerShell module

The script installs missing modules automatically if they are not already present.

### Required app permissions

The app registration must be granted the following application permissions in Azure AD / Microsoft Graph, then admin consent must be granted:

- `Exchange.ManageAsApp` (for app-only access to Exchange Online mailboxes and group mailboxes)
- `Group.Read.All` or equivalent Microsoft Graph permissions to read unified group membership and owners, if using group mailbox support

In addition, the app registration needs Azure AD admin consent so the script can authenticate app-only and query mailbox/group data without interactive sign-in.

## Configuration

Create or update `connection.json` with your tenant and app credentials:

```json
{
  "TenantId": "your-tenant-id.onmicrosoft.com",
  "AppId": "your-app-registration-client-id",
  "AppSecret": "your-app-secret",
  "UseCertificate": false,
  "CertificateThumbprint": null
}
```

If using certificate auth, set `UseCertificate` to `true` and provide `CertificateThumbprint`.

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
