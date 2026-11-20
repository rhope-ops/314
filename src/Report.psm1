<#
Report.psm1 - console summary table and CSV report for a PII scan.

  Purpose:     Builds report rows, prints a human-readable summary, and saves one CSV.
  Network:     none.
  Writes:      one CSV file in the output folder the caller gives (default: Documents\PII-Scan).
               Values in it are masked by PiiDetectors.psm1; no full PII is written.
  Credentials: none.
  Runs on:     Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version 2.0

function New-ReportRow {
    <#
    .SYNOPSIS  One CSV row: a hit, or a 'NOT SCANNED' / 'BLOCKED' note for an item.
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source      = $Source
        Channel     = $Channel
        Item        = $Item
        Detector    = $Detector
        Confidence  = $Confidence
        MaskedValue = $MaskedValue
        Context     = $Context
        Modified    = $Modified
        Link        = $Link
    }
}

function Get-ReportSummary {
    <#
    .SYNOPSIS  Groups hit rows per item: count, detectors found, highest confidence.
    .OUTPUTS   Objects sorted with the highest confidence and most hits first.
    #>
    param([object[]]$Rows)
    $rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }
    $hits = @($Rows | Where-Object { $rank.ContainsKey($_.Confidence) })
    $groups = $hits | Group-Object Source, Channel, Item
    $summary = foreach ($g in $groups) {
        $first = $g.Group[0]
        $best = ($g.Group | Sort-Object { $rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Highest   = $best
            Hits      = $g.Count
            Channel   = $first.Channel
            Source    = $first.Source
            Item      = $first.Item
            Detectors = (($g.Group | Select-Object -ExpandProperty Detector -Unique) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { $rank[$_.Highest] }; Descending = $true }, @{ Expression = 'Hits'; Descending = $true })
}

function Write-ScanSummary {
    <#
    .SYNOPSIS  Prints the summary table, the totals per detector and the items that were not scanned.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [string]$CsvPath)
    $rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }
    $hits = @($Rows | Where-Object { $rank.ContainsKey($_.Confidence) })
    $notes = @($Rows | Where-Object { -not $rank.ContainsKey($_.Confidence) })

    Write-Host ''
    Write-Host '==== PII scan summary ====' -ForegroundColor Cyan
    Write-Host ("Items scanned: {0}   Items with possible PII: {1}   Possible PII hits: {2}   Not scanned / blocked: {3}" -f `
            $ItemsScanned, @(Get-ReportSummary $Rows).Count, $hits.Count, $notes.Count)
    if ($hits.Count -gt 0) {
        Write-Host ''
        Write-Host 'Items with possible PII (highest confidence first):' -ForegroundColor Yellow
        Get-ReportSummary $Rows | Format-Table Highest, Hits, Channel, Source, Item, Detectors -AutoSize -Wrap | Out-String -Width 200 | Write-Host
        Write-Host 'Hits per type:' -ForegroundColor Yellow
        # Plain objects: Select-Object on a GroupInfo keeps its built-in view, which prints nothing here.
        $hits | Group-Object Detector, Confidence | Sort-Object Name |
            ForEach-Object { [pscustomobject]@{ Type = $_.Group[0].Detector; Confidence = $_.Group[0].Confidence; Hits = $_.Count } } |
            Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    }
    else {
        Write-Host 'No possible PII found in the items that were scanned.' -ForegroundColor Green
    }
    if ($notes.Count -gt 0) {
        Write-Host 'Not scanned or blocked (check these by hand):' -ForegroundColor Yellow
        $notes | Format-Table Detector, Channel, Item, Context -AutoSize -Wrap | Out-String -Width 200 | Write-Host
    }
    Write-Host "Full report (values masked): $CsvPath" -ForegroundColor Cyan
    Write-Host 'Confidence is a pattern guess, not proof: open each item and check it before you act.'
}

function Export-ScanCsv {
    <#
    .SYNOPSIS  Saves all rows to a new time-stamped CSV in the output folder and returns its path.
    #>
    param([object[]]$Rows, [string]$OutputFolder)
    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    # Time-stamped name; add -2, -3 ... if a report with that name already exists (never overwrite).
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $path = Join-Path $OutputFolder ('PIIScan-{0}.csv' -f $stamp)
    $i = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $OutputFolder ('PIIScan-{0}-{1}.csv' -f $stamp, $i); $i++ }
    $cols = 'Source', 'Channel', 'Item', 'Detector', 'Confidence', 'MaskedValue', 'Context', 'Modified', 'Link'
    if (@($Rows).Count -eq 0) {
        # Header only, so an empty result is still a file the user can open.
        ($cols | ForEach-Object { '"' + $_ + '"' }) -join ',' | Set-Content -LiteralPath $path -Encoding UTF8
    }
    else {
        $Rows | Select-Object $cols | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    }
    return $path
}

Export-ModuleMember -Function New-ReportRow, Get-ReportSummary, Write-ScanSummary, Export-ScanCsv
