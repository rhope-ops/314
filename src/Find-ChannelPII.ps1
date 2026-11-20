<#
.SYNOPSIS
  Find-ChannelPII.ps1 - DoD365 PII Scan: read-only check of one Teams team or SharePoint site for possible PII.

.DESCRIPTION
  Scans the files of a team's or site's "Shared Documents" library (every folder: in a team, one per standard
  channel; and the files at the top of the library), then tries to scan a team's channel posts and replies.
  Two progress bars (the whole scan, and the file in work) and one line per folder show the counts as it goes.
  Shows a readable summary on screen and saves a readable text report and a CSV. Every value found is MASKED
  on screen and in the files.

  Purpose:     Help a team member (owner rights are not needed) find files and posts that may hold PII, so they can review them.
               A spot check, not a replacement for Microsoft Purview DLP or Content Search.
  Inputs:      the team's Teams link or site address (asked when not given).
  Outputs:     progress and the summary on screen; PIIScan-<time>.txt and PIIScan-<time>.csv in -OutputFolder.
  How to run:  .\Find-ChannelPII.ps1   (README.md "Run it" gives every step)
  Run target:  Windows PowerShell 5.1 (also PowerShell 7) with PnP.PowerShell 1.12.0. No other modules.
  Network:     only the one team's SharePoint site, then Microsoft Graph for that one team (posts).
  Writes:      only the report files in -OutputFolder (default Documents\PII-Scan); nothing in Teams or SharePoint.
  Credentials: the person's own CAC sign-in in the PnP sign-in window; nothing is stored or saved.

  Policy:      Read-only; one team only (no network reconnaissance); file content stays in memory; values
               masked (the PAM, 2026-11-20). No script file is collected or sent: script types are not
               read and only their (masked) names appear in the report.
  Not scanned: PDF, images, legacy .doc/.xls/.ppt (reported as NOT SCANNED); private and shared channels
               (each has its own site: run again with that site's address); 1:1, group and meeting chats.

.PARAMETER SiteUrl
  The team's Teams link or its SharePoint site address.
  Teams link: in Teams, open the team or one of its channels, click the ... next to its name > Copy link
  (https://<teams-host>/l/team/...?groupId=...). The script asks for the team's site address and checks the site belongs to that team.
  Site address: the team's Shared tab > ... > Open in SharePoint (https://usaf.dps.mil/teams/<team-site>);
  a file's Copy link from the Shared tab works too. Anything after the site name is cut off.
.PARAMETER OutputFolder
  Where to save the report files. Default: Documents\PII-Scan in your profile.
.PARAMETER MaxFileMB
  Files larger than this are reported as NOT SCANNED. Default 25.

.EXAMPLE
  .\Find-ChannelPII.ps1
  Asks for the Teams link or site address, then scans every channel's files and posts.
.EXAMPLE
  .\Find-ChannelPII.ps1 -SiteUrl https://usaf.dps.mil/teams/<team-site>
#>
[CmdletBinding()]
param(
    [string]$SiteUrl,
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PII-Scan'),
    # 25 MB: larger Office files are rare in a channel and slow to read in memory. The author's choice, 2026-11-20.
    [int]$MaxFileMB = 25
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Modules, in dependency order (each module's header names what it needs).
foreach ($module in @('PiiDetectors.psm1', 'PiiMasking.psm1', 'TextExtract.psm1', 'Report.psm1', 'ScanMeter.psm1', 'ChannelScan.psm1')) {
    Import-Module (Join-Path $PSScriptRoot $module) -Force
}

# ==== Body ==== (scripts/build_paste.py copies everything from this line into the paste bundle; the lines above
# become the bundle's Settings block and its inline module code)

Write-Host 'DoD365 PII Scan (read-only). Values in the report are masked.' -ForegroundColor Cyan

# ---- Step 1: the Teams link or site address (only this one team is touched) ----
if (-not $SiteUrl) { $SiteUrl = Read-Host 'Teams link, or SharePoint site address (https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>)' }
$target = Get-ScanTarget $SiteUrl

# ---- Step 2: sign in with the person's own CAC (a Teams link first finds the team's site) ----
$SiteUrl = Connect-TeamSite $target

$rows = New-Object System.Collections.Generic.List[object]
$scanned = 0
$selected = @()
$finished = $false   # set only after the last step; a cancelled run (Ctrl+C) skips 'catch' but not 'finally'
try {
    # ---- Step 3: the channels to scan: all of them ----
    $library = Get-TeamDocumentLibrary
    $selected = @(Get-ChannelFolderNames $library)
    $shownFolders = ($selected | ForEach-Object { Protect-PiiText $_ }) -join ', '
    Write-Host ('Folders (channels) to scan: ' + $shownFolders + '; also any files at the top of the library')

    # ---- Step 4: the files ----
    $scanned += Invoke-FileScan -Library $library -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows

    # ---- Step 5: the posts (a refusal becomes a BLOCKED row; the file results stay) ----
    $scanned += Invoke-PostScan -SiteUrl $SiteUrl -Channels $selected -Rows $rows
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
    Add-IncompleteRow -Rows $rows -Finished $finished
    $postsLine = 'tried (see BLOCKED rows if refused)'
    $header = @(
        ('Site:     ' + $SiteUrl),
        ('Folders:  ' + (($selected | ForEach-Object { Protect-PiiText $_ }) -join ', ')),
        ('Posts:    ' + $postsLine),
        ('Run:      ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    )
    try {
        $null = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $scanned -Header $header -OutputFolder $OutputFolder
    }
    finally {
        # Sign out even when the report could not be saved (for example, a folder that cannot be written).
        try { Disconnect-PnPOnline }
        catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error for the person
    }
}
