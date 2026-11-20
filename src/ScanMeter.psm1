<#
ScanMeter.psm1 - the progress bar text and the per-channel lines of a scan: counts only, so the screen shows no PII.

  Purpose:     Lets the person follow a scan: files (or posts) done of the total, read, not scanned, possible PII,
               time used and time left (the PAM, 2026-11-20). No file name and no found value is shown; channel
               names are masked.
  Inputs:      the report row list and the counts the scan functions of ChannelScan.psm1 give.
  Outputs:     a meter (hashtable), status text, percent, and one coloured line per channel on screen.
  Depends on:  PiiMasking.psm1 (Protect-PiiText), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (screen only).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

function New-ScanMeter {
    <#
    .SYNOPSIS
      Starts a meter for one scan step: items done, read, not scanned, possible PII, and the time.
    .PARAMETER Total
      How many items the step has (0 when not known up front, as for posts).
    .OUTPUTS
      [hashtable] the meter (Update-ScanMeter and Get-ScanMeterText use it).
    #>
    param([int]$Total)
    @{ Total = $Total; Done = 0; Read = 0; Skipped = 0; Hits = 0; HitItems = 0; Clock = [Diagnostics.Stopwatch]::StartNew() }
}

function Update-ScanMeter {
    <#
    .SYNOPSIS
      Counts one item into the meter from the report rows it added.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Rows
      The report row list.
    .PARAMETER FirstRow
      The row count before the item was scanned.
    .PARAMETER ReadCount
      The items whose text was read (the scan function's return value).
    #>
    param([hashtable]$Meter, $Rows, [int]$FirstRow, [int]$ReadCount)
    $new = @(for ($r = $FirstRow; $r -lt $Rows.Count; $r++) { $Rows[$r] })
    $hits = @($new | Where-Object { $_.Confidence })   # a row with a confidence is a possible PII hit
    $Meter.Done++
    $Meter.Read += $ReadCount
    $Meter.Skipped += @($new | Where-Object { $_.Detector -eq 'NOT SCANNED' -or $_.Detector -eq 'BLOCKED' }).Count
    $Meter.Hits += $hits.Count
    $Meter.HitItems += @($hits | Select-Object -ExpandProperty Item -Unique).Count   # files or posts with a hit
}

function Get-ScanMeterText {
    <#
    .SYNOPSIS
      The meter as one status line: counts, time used and, when the total is known, time left.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is: 'file' or 'post'.
    .OUTPUTS
      [string] for example "12 of 40 files done | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 used, about 01:48 left".
    #>
    param([hashtable]$Meter, [string]$Noun)
    $elapsed = $Meter.Clock.Elapsed
    $text = [string]$Meter.Done
    if ($Meter.Total -gt 0) { $text += ' of ' + $Meter.Total }
    $text += ' ' + $Noun + 's done | read {0} | not scanned {1} | possible PII {2} in {3} {4}s | {5:mm\:ss} used' -f
        $Meter.Read, $Meter.Skipped, $Meter.Hits, $Meter.HitItems, $Noun, $elapsed
    if ($Meter.Total -gt 0 -and $Meter.Done -gt 0 -and $Meter.Done -lt $Meter.Total) {
        $left = [TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done)))
        $text += ', about {0:mm\:ss} left' -f $left
    }
    return $text
}

function Get-MeterPercent {
    # The meter's percent done (0-100); 0 when the total is not known.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0) { return 0 }
    return [int][Math]::Min(100, 100 * $Meter.Done / $Meter.Total)
}

function Write-ChannelLine {
    <#
    .SYNOPSIS
      Prints one line when a channel is done: its counts, yellow when it has possible PII.
    .PARAMETER Number
      Which channel this is (1-based) and the channel count, as "2/5".
    .PARAMETER Channel
      The channel name (masked here).
    .PARAMETER Before
      A copy of the meter before the channel (Done, Read, Skipped, Hits).
    .PARAMETER Meter
      The meter after the channel.
    .PARAMETER Noun
      What an item is: 'file' or 'post'.
    #>
    param([string]$Number, [string]$Channel, [hashtable]$Before, [hashtable]$Meter, [string]$Noun)
    $hits = $Meter.Hits - $Before.Hits
    $line = '  [{0}] {1}: {2} {3}s - read {4}, not scanned {5}, possible PII {6}' -f $Number, (Protect-PiiText $Channel),
        ($Meter.Done - $Before.Done), $Noun, ($Meter.Read - $Before.Read), ($Meter.Skipped - $Before.Skipped), $hits
    $colour = 'Green'
    if ($hits -gt 0) { $colour = 'Yellow' }
    Write-Host $line -ForegroundColor $colour
}

Export-ModuleMember -Function New-ScanMeter, Update-ScanMeter, Get-ScanMeterText, Get-MeterPercent, Write-ChannelLine
