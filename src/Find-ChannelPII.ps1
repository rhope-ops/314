<#
.SYNOPSIS
  Find-ChannelPII.ps1 - read-only spot check of one Microsoft Teams team for possible PII.

.DESCRIPTION
  Scans the files of a team's standard channels (the team site's "Shared Documents" library) and,
  with -IncludePosts, tries to scan the channel posts and replies. Shows a readable summary on screen
  and saves a readable text report and/or a CSV. Every value found is MASKED on screen and in the files.

  Purpose:     Help a team owner find files and posts that may hold PII, so they can review them.
               A spot check, not a replacement for Microsoft Purview DLP or Content Search.
  Inputs:      the team's site address, the channels, the report choice (asked when not given).
  Outputs:     the summary on screen; PIIScan-<time>.txt and/or PIIScan-<time>.csv in -OutputFolder.
  How to run:  .\Find-ChannelPII.ps1   (README.md "Run it" gives every step)
  Run target:  Windows PowerShell 5.1 (also PowerShell 7) with PnP.PowerShell 1.12.0. No other modules.
  Network:     only the one SharePoint site address given, and with -IncludePosts Microsoft Graph for that one team.
  Writes:      only the report files in -OutputFolder (default Documents\PII-Scan); nothing in Teams or SharePoint.
  Credentials: the person's own CAC sign-in in the PnP sign-in window; nothing is stored or saved.

  Policy:      Read-only; one site only (no network reconnaissance); file content stays in memory; values
               masked (the PFM, 2026-11-20). No script file is collected or sent: script types are not
               read and only their (masked) names appear in the report.
  Not scanned: PDF, images, legacy .doc/.xls/.ppt (reported as NOT SCANNED); private and shared channels
               (each has its own site: run again with that site's address); 1:1, group and meeting chats.

.PARAMETER SiteUrl
  The team's SharePoint site, for example https://<tenant>.sharepoint-mil.us/sites/<team-site>.
  In Teams: open the channel's Files tab > "Open in SharePoint" and copy the address up to the site name.
.PARAMETER Channel
  A channel folder name, or 'All'. Leave it out to pick from a list.
.PARAMETER IncludePosts
  Also try to scan channel posts and replies. Needs a second sign-in and a Graph permission that a
  tenant admin may not have granted; if it is refused, the report says BLOCKED and the file scan still counts.
.PARAMETER Report
  What to save: 'Both' (readable .txt and .csv), 'Text' (.txt only) or 'Csv' (.csv only). The summary is
  always shown on screen. Leave it out to choose from a list (Enter = Both).
.PARAMETER OutputFolder
  Where to save the report files. Default: Documents\PII-Scan in your profile.
.PARAMETER MaxFileMB
  Files larger than this are reported as NOT SCANNED. Default 25.

.EXAMPLE
  .\Find-ChannelPII.ps1
  Asks for the site address, the channels and the report choice, then scans the files.
.EXAMPLE
  .\Find-ChannelPII.ps1 -SiteUrl https://<tenant>.sharepoint-mil.us/sites/<team-site> -Channel All -IncludePosts -Report Both
#>
[CmdletBinding()]
param(
    [string]$SiteUrl,
    [string]$Channel,
    [switch]$IncludePosts,
    [ValidateSet('Both', 'Text', 'Csv')]
    [string]$Report,
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PII-Scan'),
    # 25 MB: larger Office files are rare in a channel and slow to read in memory. The author's choice, 2026-11-20.
    [int]$MaxFileMB = 25
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Modules, in dependency order (each module's header names what it needs).
foreach ($module in @('PiiDetectors.psm1', 'PiiMasking.psm1', 'TextExtract.psm1', 'Report.psm1', 'ChannelScan.psm1')) {
    Import-Module (Join-Path $PSScriptRoot $module) -Force
}

Write-Host 'Teams channel PII spot check (read-only). Values in the report are masked.' -ForegroundColor Cyan

# ---- Step 1: the site address (only this one site is touched) and the report choice ----
if (-not $SiteUrl) { $SiteUrl = Read-Host 'Team SharePoint site address (https://.../sites/<name>)' }
$SiteUrl = Test-TeamSiteUrl $SiteUrl
$format = Select-ReportFormat $Report

# ---- Step 2: sign in with the person's own CAC ----
Write-Host "Signing in to $SiteUrl ..."
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

$rows = New-Object System.Collections.Generic.List[object]
$scanned = 0
$selected = @()
$finished = $false   # set only after the last step; a cancelled run (Ctrl+C) skips 'catch' but not 'finally'
try {
    # ---- Step 3: the channels to scan ----
    $library = Get-TeamDocumentLibrary
    $selected = @(Select-ChannelFolders -Available @(Get-ChannelFolderNames $library) -Requested $Channel)
    Write-Host ('Channels to scan: ' + (($selected | ForEach-Object { Protect-PiiText $_ }) -join ', '))

    # ---- Step 4: the files ----
    $scanned += Invoke-FileScan -Library $library -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows

    # ---- Step 5 (optional): the posts ----
    if ($IncludePosts) { $scanned += Invoke-PostScan -SiteUrl $SiteUrl -Channels $selected -Rows $rows }
    $finished = $true
}
catch {
    # Keep what was found so far: the report says the scan stopped early, and why (PII in the message masked).
    $reason = Protect-PiiText $_.Exception.Message
    Write-Host ('The scan stopped early: ' + $reason) -ForegroundColor Red
    $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
}
finally {
    # ---- Step 6: the report (also after an error or a cancel), then sign out ----
    if (-not $finished -and -not @($rows | Where-Object { $_.Detector -eq 'INCOMPLETE' })) {
        $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
    }
    $postsLine = 'not scanned (-IncludePosts not used)'
    if ($IncludePosts) { $postsLine = 'tried (see BLOCKED rows if refused)' }
    $header = @(
        ('Site:     ' + $SiteUrl),
        ('Channels: ' + (($selected | ForEach-Object { Protect-PiiText $_ }) -join ', ')),
        ('Posts:    ' + $postsLine),
        ('Run:      ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    )
    try {
        $null = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $scanned -Header $header -Format $format -OutputFolder $OutputFolder
    }
    finally {
        # Sign out even when the report could not be saved (for example, a folder that cannot be written).
        try { Disconnect-PnPOnline }
        catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error for the person
    }
}
