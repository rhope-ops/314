<#
ScanMeter.psm1 - the progress bars and the per-folder lines of a scan: counts only, so the screen shows no PII.

  Purpose:     Lets the person follow a scan: files done of the total, read, not scanned, possible PII,
               time used and time left (the PAM, 2026-11-20). No file name and no found value is shown; channel
               names are masked.
  Inputs:      the report row list and the counts the scan functions of ChannelScan.psm1 give.
  Outputs:     a meter (hashtable), status text, percent, a per-file progress bar and one coloured line per folder.
  Depends on:  PiiMasking.psm1 (Protect-PiiText), Report.psm1 (Write-Ui), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (screen only).
  Credentials: none.
#>

Set-StrictMode -Version 2.0

# From this size on, the per-file bar says the step can take a minute (5 MB: a slow download on a gov network
# takes tens of seconds; the author's choice, 2026-11-20).
$script:LargeFileBytes = 5MB

# The per-file bar is shown only from this size on: a small file is done in a moment, and each bar update costs time in
# Windows PowerShell 5.1 (the author's choice, 2026-11-20: faster scans of many small files).
$script:FileBarBytes = 1MB

# A coloured line of live counts is printed under the progress bar at most this often (and after each folder).
$script:LiveLineSeconds = 10

# Short PII type names for the live line, in the PAM's order (Report.psm1 $SeverityOrder).
$script:ShortTypeNames = [ordered]@{ 'SSN' = 'SSN'; 'Date of birth' = 'DOB'; 'Passport' = 'Passport'; 'Payment card' = 'Card'
    'DoD ID' = 'DoD ID'; 'Home address' = 'Address'; 'Name' = 'Name'; 'Phone' = 'Phone'; 'Personal e-mail' = 'E-mail' }

function New-ScanMeter {
    <#
    .SYNOPSIS
      Starts a meter for one scan step: items done, read, not scanned, possible PII, and the time.
    .PARAMETER Total
      How many items the step has (0 when not known up front).
    .OUTPUTS
      [hashtable] the meter (Update-ScanMeter and Get-ScanMeterText use it).
    #>
    param([int]$Total)
    # Clock: anything with an Elapsed [TimeSpan] (a running stopwatch here; a fixed time in the tests).
    @{ Total = $Total; Done = 0; Read = 0; Skipped = 0; Hits = 0; HitItems = 0; Types = @{}; LastLine = [TimeSpan]::Zero
       Clock = [Diagnostics.Stopwatch]::StartNew() }
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
    $Meter.HitItems += @($hits | Select-Object -ExpandProperty Item -Unique).Count   # files with a hit
    foreach ($hit in $hits) { $Meter.Types[$hit.Detector] = 1 + [int]$Meter.Types[$hit.Detector] }
}

function Format-Duration {
    # A time as minutes:seconds; the minutes go past 59 (65 minutes = 65:00), so a long scan never wraps to 05:00.
    param([TimeSpan]$Time)
    return '{0:00}:{1:00}' -f [int][Math]::Floor($Time.TotalMinutes), $Time.Seconds
}

function Get-ScanMeterText {
    <#
    .SYNOPSIS
      The meter as one status line: counts, time used and, when the total is known, time left.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .OUTPUTS
      [string] for example "12 of 40 files done | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 used, about 01:48 left".
    #>
    param([hashtable]$Meter, [string]$Noun)
    $elapsed = $Meter.Clock.Elapsed
    $text = [string]$Meter.Done
    if ($Meter.Total -gt 0) { $text += ' of ' + $Meter.Total }
    $types = Get-TypeCountText $Meter
    if ($types) { $types = ' (' + $types + ')' }
    $text += ' ' + $Noun + 's done | read {0} | not scanned {1} | possible PII {2} in {3} {4}s{5} | {6} used' -f
        $Meter.Read, $Meter.Skipped, $Meter.Hits, $Meter.HitItems, $Noun, $types, (Format-Duration $elapsed)
    if ($Meter.Total -gt 0 -and $Meter.Done -gt 0 -and $Meter.Done -lt $Meter.Total) {
        $left = [TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done)))
        $text += ', about ' + (Format-Duration $left) + ' left'
    }
    return $text
}

function Get-TypeCountText {
    # The hits per type so far, worst type first, as "SSN 2, DOB 1"; empty when there is none.
    param([hashtable]$Meter)
    if (-not $Meter.ContainsKey('Types') -or $Meter.Types.Count -eq 0) { return '' }
    $parts = foreach ($type in $script:ShortTypeNames.Keys) {
        if ($Meter.Types[$type]) { '{0} {1}' -f $script:ShortTypeNames[$type], $Meter.Types[$type] }
    }
    return (@($parts) -join ', ')
}

function Write-MeterLine {
    <#
    .SYNOPSIS
      Prints the live counts as one coloured line under the progress bar (FULCRUM colours, Report.psm1 Write-Ui):
      done of the total and percent, read, not scanned, possible PII per type, time used and left. At most every
      $LiveLineSeconds, unless -Force. Counts only: no file name and no value.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .PARAMETER Force
      Print now (after a folder).
    #>
    param([hashtable]$Meter, [string]$Noun, [switch]$Force)
    $elapsed = $Meter.Clock.Elapsed
    if (-not $Force -and ($elapsed - $Meter.LastLine).TotalSeconds -lt $script:LiveLineSeconds) { return }
    $Meter.LastLine = $elapsed
    Write-Ui ('  > {0}/{1} {2}s {3,3}%' -f $Meter.Done, $Meter.Total, $Noun, (Get-MeterPercent $Meter)) 'Title' -NoNewline
    Write-Ui (' | read {0}' -f $Meter.Read) 'Text' -NoNewline
    $kind = 'Muted'; if ($Meter.Skipped -gt 0) { $kind = 'Accent' }
    Write-Ui (' | not scanned {0}' -f $Meter.Skipped) $kind -NoNewline
    $kind = 'Muted'; if ($Meter.Hits -gt 0) { $kind = 'Alert' }
    $types = Get-TypeCountText $Meter
    if ($types) { $types = ' (' + $types + ')' }
    Write-Ui (' | possible PII {0}{1}' -f $Meter.Hits, $types) $kind -NoNewline
    $left = ''
    if ($Meter.Total -gt 0 -and $Meter.Done -gt 0 -and $Meter.Done -lt $Meter.Total) {
        $left = ', about ' + (Format-Duration ([TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done))))) + ' left'
    }
    Write-Ui (' | {0} used{1}' -f (Format-Duration $elapsed), $left) 'Muted'
}

function Get-MeterPercent {
    # The meter's percent done (0-100); 0 when the total is not known.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0) { return 0 }
    return [int][Math]::Min(100, 100 * $Meter.Done / $Meter.Total)
}

function Write-FileStep {
    <#
    .SYNOPSIS
      A second progress bar for the file in work: its size and its step, so a large file does not look hung.
      PnP gives no byte count while it downloads, so the bar moves by step (download, read, check), not by byte.
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER Step
      What is being done now.
    .PARAMETER Percent
      How far through the file's steps (0-100).
    #>
    param([long]$Size, [string]$Step, [int]$Percent)
    if ($Size -lt $script:FileBarBytes) { return }
    $status = '{0} - {1:N1} MB' -f $Step, ($Size / 1MB)
    if ($Size -ge $script:LargeFileBytes) { $status += ' (a large file: this step can take a minute)' }
    Write-Progress -Id 2 -ParentId 1 -Activity 'This file' -Status $status -PercentComplete $Percent
}

function Write-ChannelLine {
    <#
    .SYNOPSIS
      Prints one line when a folder (a channel) is done: its counts in FULCRUM colours; possible PII in red.
    .PARAMETER Number
      Which folder this is (1-based) and the folder count, as "2/5".
    .PARAMETER Channel
      The channel name (masked here).
    .PARAMETER Before
      A copy of the meter before the channel (Done, Read, Skipped, Hits).
    .PARAMETER Meter
      The meter after the channel.
    .PARAMETER Noun
      What an item is, for example 'file'.
    #>
    param([string]$Number, [string]$Channel, [hashtable]$Before, [hashtable]$Meter, [string]$Noun)
    $hits = $Meter.Hits - $Before.Hits
    $skipped = $Meter.Skipped - $Before.Skipped
    Write-Ui ('  [{0}] {1}: ' -f $Number, (Protect-PiiText $Channel)) 'Title' -NoNewline
    Write-Ui ('{0} {1}s - read {2}' -f ($Meter.Done - $Before.Done), $Noun, ($Meter.Read - $Before.Read)) 'Text' -NoNewline
    $kind = 'Muted'; if ($skipped -gt 0) { $kind = 'Accent' }
    Write-Ui (', not scanned {0}' -f $skipped) $kind -NoNewline
    $kind = 'Muted'; if ($hits -gt 0) { $kind = 'Alert' }
    Write-Ui (', possible PII {0}' -f $hits) $kind
}

Export-ModuleMember -Function New-ScanMeter, Update-ScanMeter, Get-ScanMeterText, Get-MeterPercent, Write-MeterLine, Write-FileStep, Write-ChannelLine
