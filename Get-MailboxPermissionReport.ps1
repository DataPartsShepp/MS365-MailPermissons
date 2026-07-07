<#
.SYNOPSIS
    Connects to a Microsoft tenant using credentials from a JSON config and exports mailbox permission details to Excel.

.DESCRIPTION
    This script reads service principal sign-in details from a JSON file and a list of mailboxes/group addresses from a text or CSV file.
    It connects to Exchange Online, collects mailbox-level and inbox folder permission entries, plus Microsoft 365 group membership and owner details, and writes the results to an Excel workbook.

.PARAMETER ConfigPath
    Path to the JSON configuration file containing TenantId, AppId, AppSecret, and optional certificate details.

.PARAMETER EmailListPath
    Path to the email list file. Supports CSV with a header column named Email or plain text list of email addresses.

.PARAMETER OutputPath
    Path to the output Excel workbook. Defaults to MailboxPermissionsReport.xlsx in the current folder.

.EXAMPLE
    .\Get-MailboxPermissionReport.ps1 -ConfigPath .\connection.json -EmailListPath .\mailboxes.txt -OutputPath .\MailboxPermissionsReport.xlsx
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailListPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "MailboxPermissionsReport.xlsx"
)

function Ensure-PSGallery {
    if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
        throw "Install-Module is not available in this PowerShell session. Ensure PowerShellGet is installed and try again."
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}

    $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if (-not $psGallery) {
        Write-Host 'Registering PSGallery repository...' -ForegroundColor Yellow
        Register-PSRepository -Default -ErrorAction Stop | Out-Null
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction Stop
    }

    if ($psGallery.InstallationPolicy -ne 'Trusted') {
        Write-Host 'Setting PSGallery installation policy to Trusted...' -ForegroundColor Yellow
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop | Out-Null
    }
}

function Ensure-Module {
    param(
        [string]$Name
    )

    Ensure-PSGallery

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing PowerShell module '$Name' from PSGallery..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Load-Config {
    param(
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        throw "The config file '$Path' does not exist."
    }

    try {
        $json = Get-Content -Path $Path -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw "Unable to read or parse the config file '$Path': $_"
    }

    foreach ($required in @('TenantId','AppId')) {
        if (-not $json.$required) {
            throw "Missing required config value '$required' in '$Path'."
        }
    }

    if (-not $json.AppSecret -and -not $json.UseCertificate) {
        throw "Config must include AppSecret or set UseCertificate to true with CertificateThumbprint."
    }

    return $json
}

function Get-ConnectExchangeOnlineSecretParameterName {
    try {
        $cmd = Get-Command Connect-ExchangeOnline -ErrorAction Stop
    } catch {
        throw "Connect-ExchangeOnline is not available. Ensure the ExchangeOnlineManagement module is loaded."
    }

    if ($cmd.Parameters.ContainsKey('ClientSecret')) {
        return 'ClientSecret'
    }

    if ($cmd.Parameters.ContainsKey('AppSecret')) {
        return 'AppSecret'
    }

    return $null
}

function Connect-ExchangeOnlineFromConfig {
    param(
        $Config
    )

    Ensure-Module -Name ExchangeOnlineManagement

    $connectParams = @{
        AppId       = $Config.AppId
        Organization = $Config.TenantId
        ShowProgress = $false
        ConnectionUri = 'https://ps.outlook.com/powershell'
    }

    if ($Config.UseCertificate -eq $true -and $Config.CertificateThumbprint) {
        $connectParams.CertificateThumbprint = $Config.CertificateThumbprint
    } elseif ($Config.AppSecret) {
        $secretParam = Get-ConnectExchangeOnlineSecretParameterName
        if (-not $secretParam) {
            Write-Host "Installed ExchangeOnlineManagement module does not support app-secret authentication. Attempting to update the module..." -ForegroundColor Yellow
            try {
                Ensure-PSGallery
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
                $secretParam = Get-ConnectExchangeOnlineSecretParameterName
            } catch {
                Write-Warning "Automatic update failed: $_"
                $secretParam = $null
            }
        }

        if ($secretParam) {
            $connectParams[$secretParam] = $Config.AppSecret
        } else {
            throw "The installed ExchangeOnlineManagement module does not support app-secret authentication. Update the module manually or use certificate auth. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber"
        }
    } else {
        throw "Config must specify AppSecret or CertificateThumbprint for app-only authentication."
    }

    Write-Host "Connecting to Exchange Online tenant '$($Config.TenantId)'..." -ForegroundColor Cyan
    Connect-ExchangeOnline @connectParams -ErrorAction Stop | Out-Null
}

function Load-EmailList {
    param(
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "The email list file '$Path' does not exist."
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.csv' {
            $rows = Import-Csv -Path $Path -ErrorAction Stop
            if (-not ($rows | Get-Member -Name Email)) {
                throw "CSV file must contain a column named 'Email'."
            }
            return $rows.Email | Where-Object { [string]::IsNullOrWhiteSpace($_) -eq $false } | Select-Object -Unique
        }
        default {
            return Get-Content -Path $Path -ErrorAction Stop | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^[#;]' } | Select-Object -Unique
        }
    }
}

function Get-MailboxPermissionReport {
    param(
        [string[]]$Emails
    )

    $mailboxPermissions = @()
    $folderPermissions = @()
    $groupPermissions = @()

    foreach ($email in $Emails) {
        $mailbox = $null
        $isGroup = $false

        try {
            $mailbox = Get-Mailbox -Identity $email -ErrorAction Stop
        } catch {
            try {
                $group = Get-UnifiedGroup -Identity $email -ErrorAction Stop
                $isGroup = $true
            } catch {
                Write-Warning "Unable to find mailbox or group for '$email'. Skipping."
                continue
            }
        }

        if (-not $isGroup) {
            $mailboxName = $mailbox.PrimarySmtpAddress.ToString()

            try {
                $permissions = Get-MailboxPermission -Identity $mailboxName -ErrorAction Stop | Where-Object {
                    $_.User -and $_.User -notmatch 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false
                }
                foreach ($entry in $permissions) {
                    $mailboxPermissions += [PSCustomObject]@{
                        Mailbox             = $mailboxName
                        User                = $entry.User.ToString()
                        AccessRights        = ($entry.AccessRights -join ', ')
                        IsInherited         = $entry.IsInherited
                        Deny                = $entry.Deny
                        InheritanceType     = $entry.InheritanceType
                        LastKnownAccessTime = $entry.LastAccessTime
                    }
                }
            } catch {
                Write-Warning "Unable to retrieve mailbox permissions for '$mailboxName'."
            }

            foreach ($folderName in @('Inbox','Calendar','Contacts','Tasks','Sent Items')) {
                $folderIdentity = "$mailboxName`:$folderName"
                try {
                    $folderEntries = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop | Where-Object { $_.User -and $_.User -ne 'Default' }
                    foreach ($folderEntry in $folderEntries) {
                        $folderPermissions += [PSCustomObject]@{
                            Mailbox      = $mailboxName
                            Folder       = $folderName
                            User         = $folderEntry.User.ToString()
                            AccessRights = ($folderEntry.AccessRights -join ', ')
                            IsDefault    = $folderEntry.User -eq 'Default'
                        }
                    }
                } catch {
                    Write-Verbose "Folder permission lookup skipped for '$folderIdentity': $_"
                }
            }
        } else {
            $groupName = $group.PrimarySmtpAddress.ToString()
            $groupPermissions += [PSCustomObject]@{
                Group             = $groupName
                DisplayName       = $group.DisplayName
                ExternalDirectoryObjectId = $group.ExternalDirectoryObjectId
                AccessType        = $group.AccessType
                HiddenFromAddressListsEnabled = $group.HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
                AutoSubscribeNewMembers = $group.AutoSubscribeNewMembers
                Members           = ''
                Owners            = ''
            }

            try {
                $members = Get-UnifiedGroupLinks -Identity $groupName -LinkType Members -ErrorAction Stop | Select-Object -ExpandProperty PrimarySmtpAddress
                $owners  = Get-UnifiedGroupLinks -Identity $groupName -LinkType Owners -ErrorAction Stop | Select-Object -ExpandProperty PrimarySmtpAddress
                $memberString = if ($members) { $members -join '; ' } else { '' }
                $ownerString  = if ($owners) { $owners -join '; ' } else { '' }
                $groupPermissions[-1].Members = $memberString
                $groupPermissions[-1].Owners  = $ownerString
            } catch {
                Write-Warning "Unable to retrieve group links for '$groupName'."
            }
        }
    }

    return @{ MailboxPermissions = $mailboxPermissions; FolderPermissions = $folderPermissions; GroupPermissions = $groupPermissions }
}

function Export-PermissionReportToExcel {
    param(
        [psobject]$Report,
        [string]$Path
    )

    Ensure-Module -Name ImportExcel

    $directory = [System.IO.Path]::GetDirectoryName((Resolve-Path -Path $Path).Path)
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    Write-Host "Writing results to '$Path'..." -ForegroundColor Green
    $Report.MailboxPermissions | Export-Excel -Path $Path -WorksheetName 'MailboxPermissions' -AutoSize -ClearSheet
    $Report.FolderPermissions | Export-Excel -Path $Path -WorksheetName 'FolderPermissions' -AutoSize -ClearSheet -Append
    $Report.GroupPermissions | Export-Excel -Path $Path -WorksheetName 'GroupPermissions' -AutoSize -ClearSheet -Append
}

try {
    $config = Load-Config -Path $ConfigPath
    $emails = Load-EmailList -Path $EmailListPath

    if (-not $emails) {
        throw "No email addresses were loaded from '$EmailListPath'."
    }

    Connect-ExchangeOnlineFromConfig -Config $config

    $report = Get-MailboxPermissionReport -Emails $emails
    Export-PermissionReportToExcel -Report $report -Path $OutputPath

    Write-Host "Permission report created successfully: $OutputPath" -ForegroundColor Cyan
} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
