<#
Report.Tests.ps1 - tests of Report.psm1 and XlsxWriter.psm1: the screen summary and the results workbook.

  Purpose:     The workbook goes to <site>\<day>, never overwrites, has the Summary tab first and one tab per PII
               type worst first, holds masked values only and no formula; the folder column is named for a team
               or a site; the screen shows hits per type.
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

# ---- The workbook: Summary first, one tab per type (worst first), Low confidence, Not scanned ----
$dob = @(Find-PiiInText 'DOB: 03/14/1985')[0]
$reportRows += (New-ReportRow 'File' 'Ops' 'Recall.xlsx' '' '' $dob.Detector $dob.Confidence $dob.MaskedValue $dob.Context)
$header = [ordered]@{ Site = 'https://tenant.example.com/teams/TeamA'; Channels = 'General, Ops' }
$path = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site -FolderLabel 'Channel' 6>$null
$day = Split-Path $path -Parent
Assert-Equal ('TeamA/' + (Get-Date -Format 'yyyy-MM-dd')) ((Split-Path (Split-Path $day -Parent) -Leaf) + '/' + (Split-Path $day -Leaf)) 'Saved in <site>\<day>'
Assert-Equal $true ((Split-Path $path -Leaf) -match '^Ion Results \d\d-\d\d-\d\d\.xlsx$') 'Named Ion Results <time>.xlsx'
$book = Get-XlsxSheets $path
Assert-Equal 'Summary|SSN|Date of birth|Low confidence|Not scanned' (@($book.Keys) -join '|') 'Tabs: Summary, types worst first, Low, Not scanned'
Assert-Equal 'Confidence|Channel|Item|Masked value|Context|Modified|Link|Found in' ($book['SSN'][0] -join '|') 'Hit tab header, folder column named Channel for a team'
Assert-Equal 'High|General|Roster.docx|***-**-6789' (($book['SSN'][1] | Select-Object -First 4) -join '|') 'SSN tab row'
Assert-Equal 'Phone' $book['Low confidence'][1][0] 'Low hit in the Low confidence tab, with its type'
Assert-Equal '=HYPERLINK("x")' $book['Low confidence'][1][5] 'Formula-like text kept as plain text (an inline string is never a formula)'
Assert-Equal 'NOT SCANNED|Ops|scan.pdf' (($book['Not scanned'][1] | Select-Object -First 3) -join '|') 'Not scanned tab row'
$summary = ($book['Summary'] | ForEach-Object { $_ -join '|' }) -join "`n"
Assert-Equal $true ($summary -match 'Possible PII hits\|3') 'Summary: total hits'
Assert-Equal $true ($summary -match '(?m)^SSN\|1\|1\|0\|0$' -and $summary -match '(?m)^Date of birth\|1\|1\|0\|0$' -and $summary -match '(?m)^Phone\|1\|0\|0\|1$') 'Summary: hits per type, worst first'
Assert-Equal $true ($summary -match 'potential or actual breach' -and $summary -match 'para 3\.1\.2') 'Summary gives the cited breach step'
$allCells = ($book.Values | ForEach-Object { $_ | ForEach-Object { $_ -join ' ' } }) -join ' '
Assert-Equal $false $allCells.Contains($Sample.Ssn) 'Workbook has no full SSN'
$again = Publish-ScanReport -Rows $reportRows -ItemsScanned 2 -Header $header -OutputFolder $reportFolder -SiteUrl $header.Site 6>$null
Assert-Equal $true ($again -ne $path) 'A second run never overwrites the first'
Assert-Equal 'Folder' (Get-XlsxSheets $again)['SSN'][0][1] 'Folder column named Folder for a plain site'
$plain = @(Get-WorkbookSheets @() 0 ([ordered]@{}) 'Folder')
Assert-Equal 'Summary' (@($plain | ForEach-Object { $_.Name }) -join '|') 'No hits: only the Summary tab'

# ---- The screen summary ----
$screen = (Get-ScanSummaryText -Rows $reportRows -ItemsScanned 2 -Header $header -FolderLabel 'Channel') -join "`n"
Assert-Equal $true ($screen -match 'POSSIBLE PII BY TYPE' -and $screen -match 'SSN\s+1\s+0\s+0\s+1') 'Screen: hits per type with High, Medium, Low, Total'
Assert-Equal $true ($screen.IndexOf('SSN ') -lt $screen.IndexOf('Date of birth')) 'Screen: worst type first'
Assert-Equal $false $screen.Contains($Sample.Ssn) 'Screen: no full SSN'
$masked = Get-ResultPath $reportFolder ('https://tenant.example.com/teams/' + $Sample.SsnBare)
Assert-Equal $false $masked.Contains($Sample.SsnBare) 'A site name that looks like PII is masked in the folder name'

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
