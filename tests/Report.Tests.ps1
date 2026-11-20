<#
Report.Tests.ps1 - tests of Report.psm1: the readable text report, the CSV and the report choice.

  Purpose:     The choice (Both, Text, Csv) saves the right files; reports are never overwritten; the CSV
               holds masked values only and no cell Excel would read as a formula; empty results still
               give a file.
  Inputs:      $Sample (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      report files in a new temporary folder, deleted at the end of this file.
  Credentials: none.
#>

$reportFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-report-' + [guid]::NewGuid())
$hit = @(Find-PiiInText "SSN $($Sample.Ssn)")[0]
$reportRows = @(
    (New-ReportRow 'File' 'General' 'Roster.docx' 'https://x/y' '2026-11-01' $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context),
    (New-ReportRow 'File' 'General' 'Bad.docx' '' '' 'Phone' 'Low' '***-***-0142' '=HYPERLINK("x")'),
    (New-ReportRow 'File' 'Ops' 'scan.pdf' '' '' 'NOT SCANNED' '' '' 'File type is not read by this script')
)

# ---- The choice decides which files are saved ----
$both = @(Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header @('Site: test') -Format 'Both' -OutputFolder $reportFolder 6>$null)
Assert-Equal '.csv .txt' (($both | ForEach-Object { [System.IO.Path]::GetExtension($_) } | Sort-Object) -join ' ') 'Both saves .txt and .csv'
$textOnly = @(Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header @() -Format 'Text' -OutputFolder $reportFolder 6>$null)
Assert-Equal '.txt' ([System.IO.Path]::GetExtension($textOnly[0])) 'Text saves .txt only'
$csvOnly = @(Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header @() -Format 'Csv' -OutputFolder $reportFolder 6>$null)
Assert-Equal '.csv' ([System.IO.Path]::GetExtension($csvOnly[0])) 'Csv saves .csv only'
Assert-Equal 1 $csvOnly.Count 'Csv saves one file'
Assert-Equal 4 @(Get-ChildItem -LiteralPath $reportFolder).Count 'No report overwritten'
Assert-Equal 'Both' (Select-ReportFormat 'Both') 'Given choice used without asking'
Assert-Equal 'Csv' (Select-ReportFormat 'csv') 'Choice accepted in any case'
$typoRefused = $false
try { Select-ReportFormat 'Txt' | Out-Null } catch { $typoRefused = $true }  # the refusal is the expected result
Assert-Equal $true $typoRefused 'A report choice that is not Both, Text or Csv is refused before the scan'

# ---- The readable report ----
$text = Get-Content -LiteralPath ($both | Where-Object { $_ -like '*.txt' }) -Raw
Assert-Equal $true ($text.Contains('ITEMS WITH POSSIBLE PII') -and $text.Contains('HITS PER TYPE') -and $text.Contains('EACH HIT')) 'Text report sections'
Assert-Equal $true ($text -match 'SSN\s+High\s+1') 'Hits-per-type table filled'
Assert-Equal $true $text.Contains('scan.pdf') 'Not-scanned item listed'
Assert-Equal $false $text.Contains($Sample.Ssn) 'Text report has no full SSN'
Assert-Equal $true ($text -match 'servicing Privacy Manager/Monitor without delay' -and $text -match 'para 3\.1\.2') 'Text report gives the cited breach step'

# ---- The CSV ----
$csv = Get-Content -LiteralPath ($both | Where-Object { $_ -like '*.csv' }) -Raw
Assert-Equal $true $csv.Contains('***-**-6789') 'CSV has the masked SSN'
Assert-Equal $false $csv.Contains($Sample.Ssn) 'CSV has no full SSN'
Assert-Equal $true $csv.Contains("""'=HYPERLINK") 'Formula-like cell made plain text'
$empty = Export-ScanCsv -Rows @() -OutputFolder $reportFolder -Stamp 'empty'
Assert-Equal 1 @(Get-Content -LiteralPath $empty).Count 'Empty CSV has the header row'

# ---- A run that did not reach its end (Ctrl+C runs 'finally' but not 'catch') ----
$cancelRows = New-Object System.Collections.Generic.List[object]
Add-IncompleteRow -Rows $cancelRows -Finished $false
Assert-Equal 'INCOMPLETE' $cancelRows[0].Detector 'Unfinished run gets an INCOMPLETE row'
Add-IncompleteRow -Rows $cancelRows -Finished $false
Assert-Equal 1 $cancelRows.Count 'Only one INCOMPLETE row'
$doneRows = New-Object System.Collections.Generic.List[object]
Add-IncompleteRow -Rows $doneRows -Finished $true
Assert-Equal 0 $doneRows.Count 'Finished run gets no INCOMPLETE row'

Remove-Item -LiteralPath $reportFolder -Recurse -Force
