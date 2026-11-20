<#
Report.psm1 - the scan results for a person: on screen, as a readable text file and as a CSV.

  Purpose:     Holds the report rows, builds the readable summary, asks which files to save
               (the PAM, 2026-11-20: offer a readable report and a CSV), and saves them.
  Inputs:      report rows (New-ReportRow) from ChannelScan.psm1; values in them are already masked.
  Outputs:     the summary on screen; PIIScan-<time>.txt and/or PIIScan-<time>.csv.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the .txt and/or .csv report in the output folder the caller gives (default Documents\PII-Scan).
               A new file each run; an existing report is never overwritten.
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes
# (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# CSV columns, in order.
$script:CsvColumns = @('Source', 'Channel', 'Item', 'Detector', 'Confidence', 'MaskedValue', 'Context', 'Modified', 'Link')

# Line width for the text tables. Without a console window, Out-String has no width and prints
# empty tables, so a width is always given. 220 fits the widest table on a landscape printout.
$script:TableWidth = 220

# What to do with PII found where it should not be: the breach steps of the rule source, AFI 33-332 para 3.1.2 as
# changed by AFI33-332_DAFGM2026-01 (7 July 2026; edition: docs/DECISIONS.md).
# The wording is the same on the site page (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, page 1, ' +
                       'second paragraph).'

# Report choices (Select-ReportFormat): what is saved besides the on-screen summary.
$script:ReportFormats = @('Both', 'Text', 'Csv')

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name, Post or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the channel, or the post (masked by the caller).
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
      [pscustomobject] with the CSV columns.
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

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: highest confidence, number of hits, types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Highest, Hits, Channel, Source, Item, Detectors; highest confidence and most hits first.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Source, Channel, Item)) {
        $first = $group.Group[0]
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = (($group.Group | Select-Object -ExpandProperty Detector -Unique) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true }, @{ Expression = 'Hits'; Descending = $true })
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

function Get-HitsPerType {
    <#
    .SYNOPSIS
      Counts the hits per PII type and confidence. Builds plain objects: Select-Object on a
      GroupInfo keeps its built-in display, which prints an empty table here.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, Confidence, Hits.
    #>
    param([object[]]$Hits)
    $Hits | Group-Object Detector, Confidence | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ Type = $_.Group[0].Detector; Confidence = $_.Group[0].Confidence; Hits = $_.Count }
    }
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The readable report as text lines: totals, items with possible PII, hits per type, each hit,
      and the items not scanned. The same lines are shown on screen and saved as the .txt report.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files and posts were read.
    .PARAMETER Header
      Lines to put at the top (site, channels, time).
    .OUTPUTS
      [string[]]
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [string[]]$Header)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    $items = @(Get-ReportSummary $Rows)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('==== Teams channel PII spot check ====')
    foreach ($line in $Header) { $lines.Add($line) }
    $lines.Add(('Items scanned: {0}   Items with possible PII: {1}   Possible PII hits: {2}   Not scanned / blocked: {3}' -f
            $ItemsScanned, $items.Count, $hits.Count, $notes.Count))
    if ($hits.Count -gt 0) {
        $lines.Add(''); $lines.Add('ITEMS WITH POSSIBLE PII (highest confidence first):')
        $lines.Add((ConvertTo-TableText $items @('Highest', 'Hits', 'Channel', 'Source', 'Item', 'Detectors')))
        $lines.Add(''); $lines.Add('HITS PER TYPE:')
        $lines.Add((ConvertTo-TableText @(Get-HitsPerType $hits) @('Type', 'Confidence', 'Hits')))
        $lines.Add(''); $lines.Add('EACH HIT (values masked):')
        $lines.Add((ConvertTo-TableText @($hits | Sort-Object Channel, Item) @('Confidence', 'Detector', 'Item', 'MaskedValue', 'Context')))
    }
    else {
        $lines.Add(''); $lines.Add('No possible PII found in the items that were scanned.')
    }
    if ($notes.Count -gt 0) {
        $lines.Add(''); $lines.Add('NOT SCANNED, BLOCKED OR INCOMPLETE (check these by hand):')
        $lines.Add((ConvertTo-TableText $notes @('Detector', 'Channel', 'Item', 'Context')))
    }
    $lines.Add('')
    $lines.Add('Confidence is a pattern guess, not proof: open each item and check it before you act.')
    $lines.Add('Not detected: names without a label (Name:, Spouse: ...), and text inside PDFs and images.')
    $lines.Add($script:BreachAdvice)
    return $lines.ToArray()
}

# =====================================================================================
# Saving the files
# =====================================================================================

function New-ReportPath {
    <#
    .SYNOPSIS
      A new report file path: PIIScan-<time><extension>, with -2, -3 ... added if that name exists,
      so a report is never overwritten. Creates the folder if needed.
    .PARAMETER OutputFolder
      The folder for the reports.
    .PARAMETER Stamp
      The time stamp (yyyyMMdd-HHmmss).
    .PARAMETER Extension
      '.txt' or '.csv'.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$Stamp, [string]$Extension)
    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force -ErrorAction Stop | Out-Null }
    $path = Join-Path $OutputFolder ('PIIScan-{0}{1}' -f $Stamp, $Extension)
    $copy = 2
    while (Test-Path -LiteralPath $path) {
        $path = Join-Path $OutputFolder ('PIIScan-{0}-{1}{2}' -f $Stamp, $copy, $Extension)
        $copy++
    }
    return $path
}

function ConvertTo-SafeCell {
    # Excel reads a cell that starts with = + - @, a tab or a carriage return as a formula. A leading
    # apostrophe makes it plain text, so text from a scanned file cannot plant a formula in the CSV.
    param([AllowEmptyString()][string]$Value)
    if ($Value -match '^[=+\-@\t\r]') { return "'" + $Value }
    return $Value
}

function Export-ScanCsv {
    <#
    .SYNOPSIS
      Saves all rows to a new CSV (one row per hit or note). An empty result still gets the header row.
    .PARAMETER Rows
      All report rows.
    .PARAMETER OutputFolder
      The folder for the report.
    .PARAMETER Stamp
      The time stamp shared with the .txt report of the same run.
    .OUTPUTS
      [string] the CSV path.
    #>
    param([object[]]$Rows, [string]$OutputFolder, [string]$Stamp)
    $path = New-ReportPath $OutputFolder $Stamp '.csv'
    if (@($Rows).Count -eq 0) {
        ($script:CsvColumns | ForEach-Object { '"' + $_ + '"' }) -join ',' | Set-Content -LiteralPath $path -Encoding UTF8 -ErrorAction Stop
        return $path
    }
    $safeRows = foreach ($row in $Rows) {
        $cells = [ordered]@{}
        foreach ($column in $script:CsvColumns) { $cells[$column] = ConvertTo-SafeCell ([string]$row.$column) }
        [pscustomobject]$cells
    }
    $safeRows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    return $path
}

function Export-ScanText {
    <#
    .SYNOPSIS
      Saves the readable summary lines to a new .txt file.
    .PARAMETER Lines
      Output of Get-ScanSummaryText.
    .PARAMETER OutputFolder
      The folder for the report.
    .PARAMETER Stamp
      The time stamp shared with the .csv report of the same run.
    .OUTPUTS
      [string] the .txt path.
    #>
    param([string[]]$Lines, [string]$OutputFolder, [string]$Stamp)
    $path = New-ReportPath $OutputFolder $Stamp '.txt'
    $Lines | Set-Content -LiteralPath $path -Encoding UTF8 -ErrorAction Stop
    return $path
}

function Select-ReportFormat {
    <#
    .SYNOPSIS
      Which files to save: 'Both' (readable .txt and .csv), 'Text' or 'Csv'. Uses -Report when the
      person gave it; else asks, and Enter means Both.
    .PARAMETER Requested
      The -Report value, or empty.
    .OUTPUTS
      [string] Both, Text or Csv.
    #>
    param([string]$Requested)
    if ($Requested) {
        # A typo in the paste bundle's Settings (for example 'Txt') must stop before the scan, not after it.
        $match = @($script:ReportFormats | Where-Object { $_ -ieq $Requested })
        if ($match.Count -eq 0) { throw "Not a valid report choice: '$Requested'. Use Both, Text or Csv." }
        return $match[0]
    }
    Write-Host 'Save the results as:'
    Write-Host '  1. Both - readable text report (.txt) and spreadsheet (.csv)   [Enter]'
    Write-Host '  2. Text - readable text report (.txt) only'
    Write-Host '  3. Csv  - spreadsheet (.csv) only'
    $answer = (Read-Host 'Type 1, 2 or 3').Trim()
    if ($answer -eq '') { return $script:ReportFormats[0] }
    $number = 0
    if ([int]::TryParse($answer, [ref]$number) -and $number -ge 1 -and $number -le $script:ReportFormats.Count) {
        return $script:ReportFormats[$number - 1]
    }
    throw "Not a valid choice: '$answer'. Type 1, 2 or 3."
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the files the person chose.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files and posts were read.
    .PARAMETER Header
      Lines for the top of the report.
    .PARAMETER Format
      Both, Text or Csv (Select-ReportFormat).
    .PARAMETER OutputFolder
      The folder for the reports.
    .OUTPUTS
      [string[]] the paths of the saved files.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [string[]]$Header, [string]$Format, [string]$OutputFolder)
    $lines = Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header
    $lines | ForEach-Object { Write-Host $_ }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $saved = @()
    if ($Format -eq 'Both' -or $Format -eq 'Text') { $saved += Export-ScanText -Lines $lines -OutputFolder $OutputFolder -Stamp $stamp }
    if ($Format -eq 'Both' -or $Format -eq 'Csv') { $saved += Export-ScanCsv -Rows $Rows -OutputFolder $OutputFolder -Stamp $stamp }
    Write-Host ''
    foreach ($path in $saved) { Write-Host "Saved (values masked): $path" -ForegroundColor Cyan }
    Write-Host 'Handle these files as CUI: they show where PII is.' -ForegroundColor Yellow
    return $saved
}

Export-ModuleMember -Function New-ReportRow, Add-IncompleteRow, Get-ReportSummary, Get-ScanSummaryText, Export-ScanCsv, Export-ScanText,
Select-ReportFormat, Publish-ScanReport
