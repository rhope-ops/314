<#
TestKit.ps1 - shared test tools: the assertion, the invented sample values and an Office file builder.

  Purpose:     Dot-sourced by Run-Tests.ps1 before the test files.
  Inputs:      none.
  Outputs:     $script:Pass and $script:Fail (counted by Assert-Equal); the $Sample table.
  How to run:  not on its own; run tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.

  Test data:   all invented. PII-shaped samples are built from parts when the tests run (the PFM,
               2026-11-20), so this public repo holds no complete SSN-, DoD ID- or phone-shaped string.
               Card numbers are the published test numbers of the card networks.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = '$Sample is read by the test files.')]
param()

$script:Pass = 0
$script:Fail = New-Object System.Collections.Generic.List[string]

# Invented sample values, assembled from parts.
$Sample = @{
    Ssn        = '123' + '-45-' + '6789'     # valid by the SSA rules, invented
    SsnBare    = '123' + '45' + '6789'
    SsnArea000 = '000' + '-12-' + '3456'
    SsnArea666 = '666' + '-12-' + '3456'
    SsnArea9xx = '912' + '-12-' + '3456'
    SsnGroup00 = '123' + '-00-' + '3456'
    SsnSerial0 = '123' + '-45-' + '0000'
    SsnMixed   = '123' + '-45 ' + '6789'
    DodId      = '12345' + '67890'
    Phone      = '555-201' + '-0142'
    PhoneParen = '(555) 201' + '-0142'
    Passport   = 'X12' + '34567'
    Email      = 'pat.sample' + '@example.com'
    EmailQuote = "o'connor.jane" + '@example.org'
    EmailAccent = 'jos' + [char]0xE9 + '.garc' + [char]0xED + 'a' + '@example.net'
    GovEmail   = 'first.last' + '@agency.gov'
    MilEmail   = 'first.last' + '@unit.af.' + 'mil'
    Street     = '42 Elm' + ' Street'
    CityZip    = 'Springfield, IL ' + '62704'
    Site       = 'https://tenant.example.com/sites/TeamA'
}

function Assert-Equal {
    <#
    .SYNOPSIS
      Compares two values as text (case-sensitive) and counts a pass or records a failure.
    .PARAMETER Expected
      The expected value.
    .PARAMETER Actual
      The value the code returned.
    .PARAMETER Name
      The test name, shown on a failure.
    .OUTPUTS
      None.
    #>
    param($Expected, $Actual, [string]$Name)
    if ("$Expected" -ceq "$Actual") { $script:Pass++ }
    else { $script:Fail.Add("$Name : expected [$Expected] got [$Actual]") }
}

function Get-HitSummary {
    # Find-PiiInText results as 'Detector/Confidence/MaskedValue' strings; always an array.
    param([string]$Text)
    return , @(Find-PiiInText $Text | ForEach-Object { '{0}/{1}/{2}' -f $_.Detector, $_.Confidence, $_.MaskedValue })
}

function New-ZipBytes {
    <#
    .SYNOPSIS
      Builds a minimal Office file (a zip of XML parts) in memory.
    .PARAMETER Entries
      Part name -> XML text.
    .OUTPUTS
      [byte[]]
    #>
    param([hashtable]$Entries)
    Add-Type -AssemblyName System.IO.Compression
    $memory = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($name in $Entries.Keys) {
        $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open())
        $writer.Write($Entries[$name])
        $writer.Dispose()
    }
    $zip.Dispose()
    return $memory.ToArray()
}
