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

  Test data:   all invented. PII-shaped samples are built from parts when the tests run (the PAM,
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
    EmailIdn   = 'john.q' + '@ex' + [char]0xE4 + 'mple.com'
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

function Get-XlsxSheets {
    <#
    .SYNOPSIS
      Reads a workbook back for the tests: tab name -> rows, each row an array of cell texts (inline strings
      and numbers). Only the parts XlsxWriter.psm1 writes are read.
    .PARAMETER Path
      The .xlsx file.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    $stream = [System.IO.File]::OpenRead($Path)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        # Reads one part of the zip as text.
        $read = { param($name) $r = New-Object System.IO.StreamReader($zip.GetEntry($name).Open()); try { $r.ReadToEnd() } finally { $r.Dispose() } }
        $sheets = [ordered]@{}
        $names = @([regex]::Matches((& $read 'xl/workbook.xml'), '<sheet name="([^"]*)"') | ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value) })
        for ($i = 0; $i -lt $names.Count; $i++) {
            $rows = foreach ($row in [regex]::Matches((& $read ('xl/worksheets/sheet' + ($i + 1) + '.xml')), '(?s)<row[^>]*>(.*?)</row>')) {
                , @([regex]::Matches($row.Groups[1].Value, '(?s)<c [^>]*>(?:<is><t[^>]*>(.*?)</t></is>|<v>(.*?)</v>)</c>') |
                    ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value + $_.Groups[2].Value) })
            }
            $sheets[$names[$i]] = @($rows)
        }
        return $sheets
    }
    finally { $zip.Dispose(); $stream.Dispose() }
}

function Get-ResultFiles {
    # Every results workbook under a folder (the script saves them in <site>\<day> folders).
    param([string]$Folder)
    return @(Get-ChildItem -LiteralPath $Folder -Recurse -Filter 'Ion Results *.xlsx' -ErrorAction SilentlyContinue)
}

function Get-WorkbookTypes {
    <#
    .SYNOPSIS
      The PII types and note words in a workbook, sorted, one each: the type tabs, the Type column of the Low
      confidence tab and the Status column of the Not scanned tab.
    .PARAMETER Path
      The .xlsx file.
    .OUTPUTS
      [string] the names, joined with spaces.
    #>
    param([string]$Path)
    $book = Get-XlsxSheets $Path
    $types = foreach ($tab in $book.Keys) {
        if ($tab -eq 'Low confidence' -or $tab -eq 'Not scanned') { $book[$tab] | Select-Object -Skip 1 | ForEach-Object { $_[0] } }
        elseif ($tab -ne 'Summary') { $tab }
    }
    return (@($types | Sort-Object -Unique) -join ' ')
}

