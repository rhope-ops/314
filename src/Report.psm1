<#
Report.psm1 - the scan results for a person: a summary on screen and one Excel workbook with tabs.

  Purpose:     Holds the report rows, sorts them worst PII type first, shows the summary (totals and hits per type)
               and saves the workbook: Summary tab first, one tab per type, Low confidence, Not scanned
               (the PAM, 2026-11-20: tabs to cut noise; no .txt or .csv).
  Inputs:      report rows (New-ReportRow) from ChannelScan.psm1; values in them are already masked.
  Outputs:     the summary on screen; <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx.
  Depends on:  PiiMasking.psm1 (Protect-PiiText), XlsxWriter.psm1, imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the workbook and its site and day folders under the output folder (default Documents\Ion).
               A new file each run; an existing result is never overwritten.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes
# (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# The tool's name on screen and in the workbook (the PAM, 2026-11-20).
$script:ToolTitle = 'Ion - DoD365 PII Reporter'

# PII types from most to least damaging, for sorting and for the order of the tabs. AFI 33-332 gives no ranking
# (A3.1 lists only the factors for judging harm), so this is the PAM's order (2026-11-20): SSN and date of birth
# first, then identity documents and accounts, then contact details.
$script:SeverityOrder = @('SSN', 'Date of birth', 'Passport', 'Payment card', 'DoD ID', 'Home address', 'Name', 'Phone', 'Personal e-mail')

# How many files the screen summary lists (the workbook lists all).
$script:TopItems = 10

# Line width for the text tables. Without a console window, Out-String has no width and prints
# empty tables, so a width is always given. 220 fits the widest table on a landscape printout.
$script:TableWidth = 220

# What to do with PII found where it should not be: the breach steps of the rule source, AFI 33-332 para 3.1.2 as
# changed by AFI33-332_DAFGM2026-01 (7 July 2026; edition: docs/DECISIONS.md).
# The wording is the same on the site page (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, page 1, ' +
                       'second paragraph).'

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the folder (masked by the caller).
    .PARAMETER Link
      A link to the item, or empty.
    .PARAMETER Modified
      When the item last changed, or empty.
    .PARAMETER Detector
      The PII type, or the note word.
    .PARAMETER Confidence
      High, Medium, Low, or empty for a note.
    .PARAMETER MaskedValue
      The masked value, or empty.
    .PARAMETER Context
      The masked snippet, or the reason for a note.
    .OUTPUTS
      [pscustomobject] Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link.
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source = $Source; Channel = $Channel; Item = $Item; Detector = $Detector; Confidence = $Confidence
        MaskedValue = $MaskedValue; Context = $Context; Modified = $Modified; Link = $Link
    }
}

function Add-IncompleteRow {
    <#
    .SYNOPSIS
      Adds one INCOMPLETE row when the scan did not reach its end and no INCOMPLETE row is there yet.
      Called from the script's 'finally' block, which also runs after Ctrl+C (when 'catch' does not).
    .PARAMETER Rows
      The list of report rows (changed in place).
    .PARAMETER Finished
      True when every step ran.
    .OUTPUTS
      None.
    #>
    param($Rows, [bool]$Finished)
    if ($Finished) { return }
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) { return }
    $Rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
}

function Get-SeverityIndex {
    # Where a PII type stands in $SeverityOrder (0 = worst); a type not in the list comes last.
    param([string]$Type)
    $index = [Array]::IndexOf($script:SeverityOrder, $Type)
    if ($index -lt 0) { return $script:SeverityOrder.Count }
    return $index
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: its worst type, highest confidence, number of hits and types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Worst, Highest, Hits, Channel, Source, Item, Detectors; worst type first, then highest confidence.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Channel, Item)) {
        $first = $group.Group[0]
        $worst = ($group.Group | Sort-Object { Get-SeverityIndex $_.Detector }, { -$script:Rank[$_.Confidence] } | Select-Object -First 1)
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Worst = $worst.Detector; Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = ((@($group.Group | Select-Object -ExpandProperty Detector -Unique) | Sort-Object { Get-SeverityIndex $_ }) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { Get-SeverityIndex $_.Worst } }, @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true },
        @{ Expression = 'Hits'; Descending = $true })
}

function ConvertTo-TableText {
    <#
    .SYNOPSIS
      Formats objects as a text table at a fixed width (see $TableWidth).
    .PARAMETER InputRows
      The objects to show.
    .PARAMETER Columns
      The property names to show, in order.
    .OUTPUTS
      [string]
    #>
    param([object[]]$InputRows, [string[]]$Columns)
    return ($InputRows | Format-Table -Property $Columns -AutoSize -Wrap | Out-String -Width $script:TableWidth).TrimEnd()
}

function Get-TypeCounts {
    <#
    .SYNOPSIS
      The hits per PII type, worst type first, split by confidence.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, High, Medium, Low, Total.
    #>
    param([object[]]$Hits)
    foreach ($group in ($Hits | Group-Object Detector | Sort-Object { Get-SeverityIndex $_.Name })) {
        $count = @{ High = 0; Medium = 0; Low = 0 }
        foreach ($hit in $group.Group) { $count[$hit.Confidence]++ }
        [pscustomobject]@{ Type = $group.Name; High = $count.High; Medium = $count.Medium; Low = $count.Low; Total = $group.Count }
    }
}

function Get-ScanMetrics {
    <#
    .SYNOPSIS
      The scan's totals, as label and value pairs (the screen summary and the Summary tab show the same).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([object[]]$Rows, [int]$ItemsScanned)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $metrics = [ordered]@{}
    $metrics['Files read'] = $ItemsScanned
    $metrics['Files with possible PII'] = @(Get-ReportSummary $Rows).Count
    $metrics['Possible PII hits'] = $hits.Count
    $metrics['High / Medium / Low'] = '{0} / {1} / {2}' -f @($hits | Where-Object { $_.Confidence -eq 'High' }).Count,
        @($hits | Where-Object { $_.Confidence -eq 'Medium' }).Count, @($hits | Where-Object { $_.Confidence -eq 'Low' }).Count
    $metrics['Not scanned / incomplete'] = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) }).Count
    return $metrics
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The summary for the screen: the scan's details and totals, possible PII per type (worst first), and the
      files with the most serious hits. Every hit is listed in the workbook, not here.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string[]]
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(''); $lines.Add('==== ' + $script:ToolTitle + ' ====')
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) {
        $lines.Add(('{0,-26} {1}' -f ($pair.Key + ':'), $pair.Value))
    }
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    if ($hits.Count -gt 0) {
        $lines.Add(''); $lines.Add('POSSIBLE PII BY TYPE (worst first):')
        $lines.Add((ConvertTo-TableText @(Get-TypeCounts $hits) @('Type', 'High', 'Medium', 'Low', 'Total')))
        $top = @(Get-ReportSummary $Rows | Select-Object -First $script:TopItems |
            Select-Object Worst, Highest, Hits, @{ Name = $FolderLabel; Expression = { $_.Channel } }, Item)
        $lines.Add(''); $lines.Add(('FILES WITH THE MOST SERIOUS HITS (first {0}; all are in the workbook):' -f $script:TopItems))
        $lines.Add((ConvertTo-TableText $top @('Worst', 'Highest', 'Hits', $FolderLabel, 'Item')))
    }
    else { $lines.Add(''); $lines.Add('No possible PII found in the files that were read.') }
    $lines.Add('')
    $lines.Add('Confidence is a pattern guess, not proof: open each item and check it before you act.')
    $lines.Add($script:BreachAdvice)
    return $lines.ToArray()
}

# =====================================================================================
# The workbook
# =====================================================================================

function Get-ResultPath {
    <#
    .SYNOPSIS
      A new workbook path: <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx, with (2), (3) ... added
      if the name exists, so a result is never overwritten. Creates the folders. The site name is masked like any
      name the tool shows (digits become #), and characters Windows refuses in a folder name are replaced.
    .PARAMETER OutputFolder
      The base folder.
    .PARAMETER SiteUrl
      The site address; its last part (teams/<name>) names the site folder.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$SiteUrl)
    $siteName = [uri]::UnescapeDataString(($SiteUrl.TrimEnd('/') -split '/')[-1])
    $siteName = ((Protect-PiiText $siteName) -replace '[\\/:*?"<>|]', '-').Trim(' .')
    if (-not $siteName) { $siteName = 'site' }
    if ($siteName -match '^(CON|PRN|AUX|NUL|COM\d|LPT\d)$') { $siteName = '_' + $siteName }   # names Windows keeps for devices
    $now = Get-Date
    $folder = Join-Path (Join-Path $OutputFolder $siteName) ($now.ToString('yyyy-MM-dd'))
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
    $base = 'Ion Results ' + $now.ToString('HH-mm-ss')
    $path = Join-Path $folder ($base + '.xlsx')
    $copy = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $folder ('{0} ({1}).xlsx' -f $base, $copy); $copy++ }
    return $path
}

function Get-HitCells {
    # One row of a hit tab: the hit's cells in the order of $HitColumns.
    param($Hit)
    @($Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context, $Hit.Modified, $Hit.Link, $Hit.Source)
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first, then one tab per PII type with High or Medium hits (worst type
      first), then Low confidence and Not scanned when they have rows (the PAM, 2026-11-20).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the Summary tab.
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site): the name of the folder column.
    .OUTPUTS
      [object[]] New-XlsxSheet objects.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $columns = @('Confidence', $FolderLabel, 'Item', 'Masked value', 'Context', 'Modified', 'Link', 'Found in')
    $widths = @(11, 22, 45, 18, 60, 17, 50, 10)
    $byRank = @{ Expression = { $script:Rank[$_.Confidence] }; Descending = $true }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header) @(28, 70, 9, 9, 9)))
    foreach ($type in @(Get-TypeCounts $hits | Where-Object { $_.High + $_.Medium -gt 0 } | ForEach-Object { $_.Type })) {
        $typeHits = @($hits | Where-Object { $_.Detector -eq $type -and $_.Confidence -ne 'Low' } | Sort-Object $byRank, Channel, Item)
        $sheets.Add((New-XlsxSheet $type $columns @($typeHits | ForEach-Object { , (Get-HitCells $_) }) $widths -List))
    }
    $low = @($hits | Where-Object { $_.Confidence -eq 'Low' } | Sort-Object { Get-SeverityIndex $_.Detector }, Channel, Item)
    if ($low.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Low confidence' (@('Type') + $columns) @($low | ForEach-Object { , (@($_.Detector) + (Get-HitCells $_)) }) (@(16) + $widths) -List))
    }
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    if ($notes.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'Item', 'Reason', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, $_.Link) }) @(14, 22, 45, 70, 50) -List))
    }
    return $sheets.ToArray()
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the hits per type (worst first), then how to read
      the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) { , @($pair.Key, $pair.Value) }
    , @('')
    , @('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total')
    foreach ($type in @(Get-TypeCounts $hits)) { , @($type.Type, $type.High, $type.Medium, $type.Low, $type.Total) }
    if ($hits.Count -eq 0) { , @('None found in the files that were read.') }
    , @('')
    , @('Tabs', 'One tab per type with High or Medium hits, worst type first; then Low confidence; then Not scanned.')
    , @('Confidence', 'A pattern guess, not proof: open each item and check it before you act.')
    , @('If you find PII', $script:BreachAdvice)
    , @('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.')
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the workbook (Summary tab first). Returns its path.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the summary.
    .PARAMETER OutputFolder
      The base folder; the workbook goes in <site>\<day> under it.
    .PARAMETER SiteUrl
      The site address (names the site folder).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$OutputFolder, [string]$SiteUrl,
          [string]$FolderLabel = 'Folder')
    Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header -FolderLabel $FolderLabel | ForEach-Object { Write-Host $_ }
    $path = Get-ResultPath $OutputFolder $SiteUrl
    Export-Xlsx -Path $path -Sheets (Get-WorkbookSheets $Rows $ItemsScanned $Header $FolderLabel)
    Write-Host ''
    Write-Host "Saved (values masked): $path" -ForegroundColor Cyan
    Write-Host 'Handle this file as CUI: it shows where PII is.' -ForegroundColor Yellow
    return $path
}

Export-ModuleMember -Function New-ReportRow, Add-IncompleteRow, Get-ReportSummary, Get-TypeCounts, Get-ScanSummaryText, Get-WorkbookSheets,
Get-ResultPath, Publish-ScanReport
