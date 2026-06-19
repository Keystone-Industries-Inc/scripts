<#
.SYNOPSIS
Sync an Exchange Online source distribution group to a target mail-enabled security group,
and send a Microsoft Graph email report ONLY when changes are made or when the script fails.

.DESCRIPTION
FINAL GRAPH-ONLY replacement script.
- NO SMTP relay or SMTP AUTH used anywhere.
- Uses Exchange Online app-only certificate auth for the sync.
- Uses Microsoft Graph app-only certificate auth for the email report.
- Deletes log/report files older than the configured retention period.
- Sends email ONLY when changes are made OR when the script fails.
- Does NOT email when the script succeeds with no changes.

CRITICAL RUN CONTEXT
- Run this script as the SAME Windows account that owns the certificate in Cert:\CurrentUser\My.
- In your environment, that should be the z-aadsync service account.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$ReportOnly,
    [switch]$UseBypassSecurityGroupManagerCheck
)

# =========================
# CONFIGURATION - EDIT ONLY THIS BLOCK
# =========================
$Config = [ordered]@{
    RunAsUserHint         = 'z-aadsync'
    SourceGroupIdentity   = 'GlobalManagement@keystoneind.com'
    TargetGroupIdentity   = 'Cal - Out of Office - Global Management'

    AppId                 = '1c39d553-8cdb-43ec-bcfe-a4ec0b130c69'
    TenantId              = '9e9df17f-5310-4801-a62c-049a616454a2'
    Organization          = 'polycc.onmicrosoft.com'
    CertificateThumbprint = 'F30CFBF4DD0C3C82D48A74288C9AB5A7E1D64B91'

    ReportFolder          = 'D:\Scripts\GroupSyncLogs'
    RetentionDays         = 90

    MailTo                = 'itnet@keystoneind.com'
    MailFrom              = 'group-sync@keystoneind.com'
}
# =========================
# END CONFIGURATION
# =========================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts] [$Level] $Message"
}

function ConvertTo-HtmlFragment {
    param(
        [AllowNull()]$Data,
        [Parameter(Mandatory = $true)][string]$Title
    )

    if ($null -eq $Data -or @($Data).Count -eq 0) {
        return "<h3>$Title</h3><p>None</p>"
    }

    return ($Data | ConvertTo-Html -Fragment -PreContent "<h3>$Title</h3>")
}

function Resolve-AppCertificate {
    param(
        [Parameter(Mandatory = $true)][string]$Thumbprint,
        [Parameter(Mandatory = $true)][string]$RunAsUserHint
    )

    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    Write-Log "Current Windows identity: $currentIdentity"

    if ($currentIdentity -notmatch [regex]::Escape($RunAsUserHint)) {
        throw "This script must run as the Windows account that owns the certificate in Cert:\CurrentUser\My. Expected account containing '$RunAsUserHint'. Current identity is '$currentIdentity'."
    }

    $cert = Get-ChildItem "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
    if (-not $cert) {
        throw "Certificate with thumbprint $Thumbprint was not found in Cert:\CurrentUser\My for the current account ($currentIdentity)."
    }

    if (-not $cert.HasPrivateKey) {
        throw "Certificate $Thumbprint was found but does not have a private key."
    }

    return $cert
}

function Import-RequiredModules {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'ExchangeOnlineManagement module is not installed for this account.'
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication module is not installed for this account.'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions) {
        Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    }
    elseif (Get-Module -ListAvailable -Name Microsoft.Graph.Users) {
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
    }
    else {
        throw 'Neither Microsoft.Graph.Users.Actions nor Microsoft.Graph.Users is installed for this account.'
    }

    if (-not (Get-Command Send-MgUserMail -ErrorAction SilentlyContinue)) {
        throw 'Send-MgUserMail cmdlet is not available after importing Microsoft Graph modules.'
    }
}

function New-GraphFileAttachments {
    param([string[]]$Paths)

    $attachmentObjects = @()
    foreach ($path in @($Paths | Where-Object { $_ -and (Test-Path -LiteralPath $_) })) {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $name = [System.IO.Path]::GetFileName($path)
        $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

        $contentType = switch ($ext) {
            '.csv'  { 'text/csv' }
            '.txt'  { 'text/plain' }
            '.html' { 'text/html' }
            default { 'application/octet-stream' }
        }

        $attachmentObjects += @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = $name
            contentType   = $contentType
            contentBytes  = [System.Convert]::ToBase64String($bytes)
        }
    }

    return $attachmentObjects
}

function Send-GraphReportEmail {
    param(
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$MailFrom,
        [Parameter(Mandatory = $true)][string]$MailTo,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$BodyHtml,
        [string[]]$Attachments
    )

    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $Certificate -NoWelcome
    try {
        $graphAttachments = New-GraphFileAttachments -Paths $Attachments

        $params = @{
            Message = @{
                Subject = $Subject
                Body = @{
                    ContentType = 'HTML'
                    Content     = $BodyHtml
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $MailTo
                        }
                    }
                )
            }
            SaveToSentItems = $false
        }

        if ($graphAttachments.Count -gt 0) {
            $params.Message.Attachments = $graphAttachments
        }

        Send-MgUserMail -UserId $MailFrom -BodyParameter $params
    }
    finally {
        Disconnect-MgGraph | Out-Null
    }
}

$subjectStatus = 'SUCCESS'
$hadFailure = $false
$summaryLines = New-Object System.Collections.Generic.List[string]
$attachments = New-Object System.Collections.Generic.List[string]
$sourceFile = $null
$targetBeforeFile = $null
$toAddFile = $null
$toRemoveFile = $null
$targetAfterFile = $null
$transcriptPath = $null
$cert = $null

try {
    foreach ($requiredKey in @('RunAsUserHint','SourceGroupIdentity','TargetGroupIdentity','AppId','TenantId','Organization','CertificateThumbprint','ReportFolder','RetentionDays','MailTo','MailFrom')) {
        if ([string]::IsNullOrWhiteSpace([string]$Config[$requiredKey])) {
            throw "Configuration value '$requiredKey' is blank. Edit the CONFIGURATION block at the top of the script."
        }
    }

    if (-not (Test-Path -LiteralPath $Config.ReportFolder)) {
        New-Item -ItemType Directory -Path $Config.ReportFolder -Force | Out-Null
    }

    # Delete log/report files older than the configured retention period.
    $cutoffDate = (Get-Date).AddDays(-[int]$Config.RetentionDays)
    Get-ChildItem -Path $Config.ReportFolder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoffDate } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                Write-Log ("Deleted old log/report file: {0}" -f $_.FullName)
            }
            catch {
                Write-Log ("Failed to delete old log/report file: {0}. Error: {1}" -f $_.FullName, $_.Exception.Message) 'WARN'
            }
        }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $transcriptPath = Join-Path $Config.ReportFolder ("GroupSync_Transcript_{0}.txt" -f $stamp)
    Start-Transcript -LiteralPath $transcriptPath -Force | Out-Null
    $attachments.Add($transcriptPath) | Out-Null

    Import-RequiredModules
    $cert = Resolve-AppCertificate -Thumbprint $Config.CertificateThumbprint -RunAsUserHint $Config.RunAsUserHint

    Write-Log 'Connecting to Exchange Online using certificate object authentication'
    Connect-ExchangeOnline -Certificate $cert -AppId $Config.AppId -Organization $Config.Organization -ShowBanner:$false -ErrorAction Stop

    Write-Log ("Validating source group: {0}" -f $Config.SourceGroupIdentity)
    $sourceGroup = Get-DistributionGroup -Identity $Config.SourceGroupIdentity -ErrorAction Stop

    Write-Log ("Validating target group: {0}" -f $Config.TargetGroupIdentity)
    $targetGroup = Get-DistributionGroup -Identity $Config.TargetGroupIdentity -ErrorAction Stop

    Write-Log 'Reading source group members'
    $sourceMembers = @(Get-DistributionGroupMember -Identity $sourceGroup.Identity -ResultSize Unlimited -ErrorAction Stop)

    Write-Log 'Reading target group members'
    $targetMembers = @(Get-DistributionGroupMember -Identity $targetGroup.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)

    $sourceNormalized = $sourceMembers | ForEach-Object {
        [PSCustomObject]@{
            Identity           = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString().ToLower() } else { $_.Name.ToLower() }
            DisplayName        = $_.DisplayName
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { $null }
            RecipientType      = $_.RecipientType
        }
    } | Sort-Object Identity -Unique

    $targetNormalized = $targetMembers | ForEach-Object {
        [PSCustomObject]@{
            Identity           = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString().ToLower() } else { $_.Name.ToLower() }
            DisplayName        = $_.DisplayName
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { $null }
            RecipientType      = $_.RecipientType
        }
    } | Sort-Object Identity -Unique

    $sourceIds = @($sourceNormalized | Select-Object -ExpandProperty Identity)
    $targetIds = @($targetNormalized | Select-Object -ExpandProperty Identity)

    $toAdd = @($sourceNormalized | Where-Object { $_.Identity -notin $targetIds })
    $toRemove = @($targetNormalized | Where-Object { $_.Identity -notin $sourceIds })

    $safeSource = if ($sourceGroup.Alias) { $sourceGroup.Alias } else { 'Source' }
    $safeTarget = if ($targetGroup.Alias) { $targetGroup.Alias } else { 'Target' }
    $base = Join-Path $Config.ReportFolder ("GroupSync_{0}_to_{1}_{2}" -f $safeSource, $safeTarget, $stamp)

    $sourceFile = $base + '_SourceMembers.csv'
    $targetBeforeFile = $base + '_TargetMembers_Before.csv'
    $toAddFile = $base + '_ToAdd.csv'
    $toRemoveFile = $base + '_ToRemove.csv'
    $targetAfterFile = $base + '_TargetMembers_After.csv'

    $sourceNormalized | Export-Csv -Path $sourceFile -NoTypeInformation -Encoding UTF8
    $targetNormalized | Export-Csv -Path $targetBeforeFile -NoTypeInformation -Encoding UTF8
    $toAdd | Export-Csv -Path $toAddFile -NoTypeInformation -Encoding UTF8
    $toRemove | Export-Csv -Path $toRemoveFile -NoTypeInformation -Encoding UTF8

    foreach ($file in @($sourceFile, $targetBeforeFile, $toAddFile, $toRemoveFile)) {
        $attachments.Add($file) | Out-Null
    }

    Write-Log ("Source count: {0}" -f $sourceNormalized.Count)
    Write-Log ("Target count before sync: {0}" -f $targetNormalized.Count)
    Write-Log ("Members to add: {0}" -f $toAdd.Count)
    Write-Log ("Members to remove: {0}" -f $toRemove.Count)

    $summaryLines.Add("Server: $env:COMPUTERNAME") | Out-Null
    $summaryLines.Add(("Source group: {0}" -f $Config.SourceGroupIdentity)) | Out-Null
    $summaryLines.Add(("Target group: {0}" -f $Config.TargetGroupIdentity)) | Out-Null
    $summaryLines.Add(("Source member count: {0}" -f $sourceNormalized.Count)) | Out-Null
    $summaryLines.Add(("Target member count before sync: {0}" -f $targetNormalized.Count)) | Out-Null
    $summaryLines.Add(("Members to add: {0}" -f $toAdd.Count)) | Out-Null
    $summaryLines.Add(("Members to remove: {0}" -f $toRemove.Count)) | Out-Null

    if (-not $ReportOnly) {
        $memberListForUpdate = @($sourceNormalized | ForEach-Object {
            if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress } else { $_.Identity }
        })

        if ($PSCmdlet.ShouldProcess($Config.TargetGroupIdentity, ("Replace target group membership with members from {0}" -f $Config.SourceGroupIdentity))) {
            if ($UseBypassSecurityGroupManagerCheck) {
                Update-DistributionGroupMember -Identity $targetGroup.Identity -Members $memberListForUpdate -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
            }
            else {
                Update-DistributionGroupMember -Identity $targetGroup.Identity -Members $memberListForUpdate -Confirm:$false -ErrorAction Stop
            }
            Write-Log 'Membership replacement completed successfully.'
        }

        $targetMembersAfter = @(Get-DistributionGroupMember -Identity $targetGroup.Identity -ResultSize Unlimited -ErrorAction Stop) | ForEach-Object {
            [PSCustomObject]@{
                Identity           = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString().ToLower() } else { $_.Name.ToLower() }
                DisplayName        = $_.DisplayName
                PrimarySmtpAddress = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { $null }
                RecipientType      = $_.RecipientType
            }
        } | Sort-Object Identity -Unique

        $targetMembersAfter | Export-Csv -Path $targetAfterFile -NoTypeInformation -Encoding UTF8
        $attachments.Add($targetAfterFile) | Out-Null
        $summaryLines.Add(("Target member count after sync: {0}" -f $targetMembersAfter.Count)) | Out-Null
    }
    else {
        Write-Log 'ReportOnly specified. No membership changes were made.' 'WARN'
        $summaryLines.Add('ReportOnly mode: no membership changes were made.') | Out-Null
    }
}
catch {
    $subjectStatus = 'FAILED'
    $hadFailure = $true
    Write-Log $_.Exception.Message 'ERROR'
    $summaryLines.Add(("Error: {0}" -f $_.Exception.Message)) | Out-Null
}
finally {
    try { Stop-Transcript | Out-Null } catch {}

    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }

    try {
        if (-not $cert) {
            $cert = Resolve-AppCertificate -Thumbprint $Config.CertificateThumbprint -RunAsUserHint $Config.RunAsUserHint
        }

        $addRows = @()
        $removeRows = @()

        if ($toAddFile -and (Test-Path -LiteralPath $toAddFile)) {
            $addRows = @(Import-Csv -LiteralPath $toAddFile)
        }

        if ($toRemoveFile -and (Test-Path -LiteralPath $toRemoveFile)) {
            $removeRows = @(Import-Csv -LiteralPath $toRemoveFile)
        }

        $html = @()
        $html += '<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:10pt;">'
        $html += "<h2>Global Management Group Sync Report - $subjectStatus</h2>"
        $html += '<ul>'
        foreach ($line in $summaryLines) {
            $html += "<li>$line</li>"
        }
        $html += '</ul>'
        $html += (ConvertTo-HtmlFragment -Data $addRows -Title 'Members To Add')
        $html += (ConvertTo-HtmlFragment -Data $removeRows -Title 'Members To Remove')
        $html += '</body></html>'
        $body = ($html -join [Environment]::NewLine)

        # Only send email if the script failed OR changes were made.
        $shouldSendEmail = $false
        if ($hadFailure) {
            $shouldSendEmail = $true
        }
        elseif ($addRows.Count -gt 0 -or $removeRows.Count -gt 0) {
            $shouldSendEmail = $true
        }

        if ($shouldSendEmail) {
            if ($hadFailure) {
                $subjectPrefix = '[FAILED]'
            }
            else {
                $subjectPrefix = '[SUCCESS]'
            }

            $subject = "$subjectPrefix Global Management Calendar Group Sync"

            Send-GraphReportEmail `
                -Certificate $cert `
                -ClientId $Config.AppId `
                -TenantId $Config.TenantId `
                -MailFrom $Config.MailFrom `
                -MailTo $Config.MailTo `
                -Subject $subject `
                -BodyHtml $body `
                -Attachments $attachments

            Write-Host ("Report email sent to {0} via Microsoft Graph" -f $Config.MailTo)
        }
        else {
            Write-Host 'No changes were made and no errors occurred. No email was sent.'
        }
    }
    catch {
        Write-Host ("Failed to send report email via Microsoft Graph: {0}" -f $_.Exception.Message)
        if (-not $hadFailure) {
            $hadFailure = $true
        }
    }
}

if ($hadFailure) {
    throw 'The sync job completed with one or more errors. Review the transcript and CSV reports in the report folder.'
}
