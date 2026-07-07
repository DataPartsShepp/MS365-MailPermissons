# Mail Permissions Report

This repository contains a PowerShell script to connect to Microsoft 365 via app-only authentication, read a list of mailbox addresses, and export permission details to an Excel workbook.

## Files

- `Get-MailboxPermissionReport.ps1` - main script
- `connection.json` - local app-only connection configuration (do not commit)
- `connection.example.json` - generic example configuration
- `connection.certificate.example.json` - certificate-based configuration example
- `emails.txt` - sample mailbox list

> `connection.json` and `emails.txt` are ignored by `.gitignore` to avoid committing secrets and local data.

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

> Important: `connection.json` contains sensitive credentials and should remain local. Do not commit `connection.json` to source control.

## Certificate setup

If you do not already have a certificate uploaded to your Azure AD app registration, create one and upload its public key.

### 1. Create a certificate locally

Run the following in PowerShell to create a self-signed certificate in your CurrentUser store:

```powershell
$cert = New-SelfSignedCertificate \
  -Subject "CN=MS365-MailPermissions" \
  -CertStoreLocation Cert:\CurrentUser\My \
  -KeyExportPolicy Exportable \
  -KeySpec Signature \
  -NotAfter (Get-Date).AddYears(2)

Export-Certificate -Cert $cert -FilePath .\MS365-MailPermissions.cer
```

If you see an error like `Provider type not defined`, it means your system does not support the explicit provider name. Omit `-Provider` and use the default provider instead.

This creates a certificate with a private key in `Cert:\CurrentUser\My` and exports the public certificate to `MS365-MailPermissions.cer`.

### 2. Upload the certificate to your Azure AD app

1. Open the Azure portal and go to `Azure Active Directory` > `App registrations`.
2. Select your app registration.
3. Choose `Certificates & secrets`.
4. Select `Upload certificate` and upload the `.cer` file created above.

After upload, Azure AD will display the certificate thumbprint.

### 3. Use the certificate thumbprint in `connection.json`

Use the thumbprint for the certificate that is installed in your local store and uploaded to Azure AD.

```powershell
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint, HasPrivateKey
```

Then update `connection.json`:

```json
{
  "TenantId": "your-tenant-id.onmicrosoft.com",
  "AppId": "your-app-registration-client-id",
  "AppSecret": null,
  "UseCertificate": true,
  "CertificateThumbprint": "YOUR_CERTIFICATE_THUMBPRINT"
}
```

### 4. Required app permissions

The app registration must have the `Exchange.ManageAsApp` application permission for the Office 365 Exchange Online API, and an administrator must grant consent.

## Input list

Use a plain text file or CSV file to provide mailboxes or Microsoft 365 group SMTP addresses.

- Open `certmgr.msc` and navigate to `Personal\Certificates`.
- Find the certificate you uploaded for the Azure AD app.
- View the certificate and note the `Thumbprint` value directly in the list or properties.

Or run in PowerShell:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint, HasPrivateKey
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
