<#
.SYNOPSIS
  Find-ChannelPII.ps1 - read-only spot check of one Microsoft Teams team for possible PII.

.DESCRIPTION
  Scans the files of a team's standard channels (the team site's "Shared Documents" library) and,
  with -IncludePosts, tries to scan channel posts and replies. Prints a summary table and saves a
  CSV report. All found values are MASKED in the console and the CSV.

  Purpose:     Help a team owner find files and posts that may hold PII, so they can review them.
               This is a spot check, not a replacement for Microsoft Purview DLP / Content Search.
  Network:     only the one SharePoint site URL you give, and (with -IncludePosts) Microsoft Graph
               for that one team. It never lists other sites or teams.
  Writes:      one CSV file in -OutputFolder (default: Documents\PII-Scan). It never changes, moves
               or deletes anything in Teams or SharePoint. File content is read in memory only.
  Credentials: your own CAC sign-in in the PnP web login window. Nothing is stored.
  Needs:       Windows PowerShell 5.1 (or 7) and PnP.PowerShell 1.12.0. No other modules.

  Not scanned: PDF, images, legacy .doc/.xls/.ppt (listed in the report as NOT SCANNED), private
  and shared channels (they have their own sites: run the script again with that site's URL),
  1:1 and group chats, meeting chats.

.PARAMETER SiteUrl
  The team's SharePoint site, for example https://<tenant>.sharepoint-mil.us/sites/<team-site>.
  In Teams: open the channel's Files tab > "Open in SharePoint", and copy the address up to the site name.
.PARAMETER Channel
  A channel folder name, or 'All'. Leave it out to pick from a list.
.PARAMETER IncludePosts
  Also try to scan channel posts and replies. Needs a second sign-in and a Graph permission that a
  tenant admin may not have granted; if it is blocked the report says so and the file scan still runs.
.PARAMETER OutputFolder
  Where to save the CSV. Default: Documents\PII-Scan in your profile.
.PARAMETER MaxFileMB
  Files larger than this are listed as NOT SCANNED. Default 25.

.EXAMPLE
  .\Find-ChannelPII.ps1
  Asks for the site URL and the channel, then scans the files.
.EXAMPLE
  .\Find-ChannelPII.ps1 -SiteUrl https://<tenant>.sharepoint-mil.us/sites/<team-site> -Channel All -IncludePosts
#>
[CmdletBinding()]
param(
    [string]$SiteUrl,
    [string]$Channel,
    [switch]$IncludePosts,
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PII-Scan'),
    [int]$MaxFileMB = 25
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'PiiDetectors.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'TextExtract.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Report.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'ChannelScan.psm1') -Force

Write-Host 'Teams channel PII spot check (read-only). Values in the report are masked.' -ForegroundColor Cyan

# 1. Site URL: ask if not given, then check its form so we only ever touch one site.
if (-not $SiteUrl) { $SiteUrl = Read-Host 'Team SharePoint site URL (https://.../sites/<name>)' }
$SiteUrl = Test-TeamSiteUrl $SiteUrl

# 2. Sign in with your own CAC (PnP web login window).
Write-Host "Signing in to $SiteUrl ..."
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# 3. Pick the channel folders to scan.
$library = Get-TeamDocumentLibrary
$folders = @(Get-ChannelFolderNames $library)
$selected = @(Select-ChannelFolders -Available $folders -Requested $Channel)
Write-Host ('Channels to scan: ' + ($selected -join ', '))

# 4. Scan the files.
$rows = New-Object System.Collections.Generic.List[object]
$scanned = Invoke-FileScan -Library $library -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows

# 5. Optional: try the posts.
if ($IncludePosts) {
    $scanned += Invoke-PostScan -SiteUrl $SiteUrl -Channels $selected -Rows $rows
}

# 6. Report.
$csv = Export-ScanCsv -Rows $rows.ToArray() -OutputFolder $OutputFolder
Write-ScanSummary -Rows $rows.ToArray() -ItemsScanned $scanned -CsvPath $csv
Disconnect-PnPOnline
