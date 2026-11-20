<#
EndToEnd.Tests.ps1 - runs Find-ChannelPII.ps1 from start to end against the fake site.

  Purpose:     The whole script: every folder is scanned with no question asked and one sign-in, one workbook is
               written, every planted type is found, the progress bars show the counts, a plain site gets a Folder
               column, and a failure part way still saves results that say INCOMPLETE.
  Inputs:      $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      report files in a new temporary folder, deleted at the end of this file.
  Credentials: none.
#>

$script = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'src') 'Find-ChannelPII.ps1'
$runFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-e2e-' + [guid]::NewGuid())
$library = '/sites/TeamA/Shared Documents'

# ---- A full run: one workbook, every planted type, no question, the progress bars ----
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
& $script -SiteUrl $Sample.Site -OutputFolder $runFolder 6>$null
$results = @(Get-ResultFiles $runFolder)
Assert-Equal 1 $results.Count 'Script saved one workbook'
# -Include is ignored with -LiteralPath in Windows PowerShell 5.1, so the extension is checked after the listing.
Assert-Equal 0 @(Get-ChildItem -LiteralPath $runFolder -Recurse | Where-Object { $_.Extension -in @('.txt', '.csv') }).Count 'No .txt or .csv any more'
Assert-Equal 'NOT SCANNED Phone SSN' (Get-WorkbookTypes $results[0].FullName) 'Script found every planted type'
Assert-Equal 'Channel' (Get-XlsxSheets $results[0].FullName)['SSN'][0][1] 'A team: the folder column is named Channel'
Assert-Equal $false (Test-FakePnPCall 'Read-Host') 'No question asked when the site is given'
Assert-Equal 1 @(Get-FakePnPConnection).Count 'One sign-in only: no posts sign-in'
Assert-Equal $true ((Get-FakeProgressText) -match '2 of 3 files done \| read \d+ \| not scanned 0 \| possible PII \d+ in \d+ files') 'Progress bar shows the file counts'
Assert-Equal $true ((Get-FakeProgressText) -match "folder 2 of 2: 'Ops'") 'Progress bar names the folder and its number'
Assert-Equal $true ((Get-FakeProgressText) -match 'This file Checking for PII - \d+\.\d MB') 'Per-file bar shows the step and the size'

# ---- A plain SharePoint site (no team): Folder column, top-of-library files ----
Reset-FakePnP @{ "$library/top.txt" = $Txt }
Set-FakeSiteAnswer -TypedSite '' -SiteGroupId ([guid]::Empty)
$siteFolder = Join-Path $runFolder 'site'
& $script -SiteUrl $Sample.Site -OutputFolder $siteFolder 6>$null
$siteBook = Get-XlsxSheets (Get-ResultFiles $siteFolder)[0].FullName
Assert-Equal 'Folder|(top of the library)' (($siteBook['Phone'][0][1], $siteBook['Phone'][1][1]) -join '|') 'A plain site: Folder column; top-of-library file reported'

# ---- A run that fails part way still saves a report ----
Reset-FakePnP @{}
Set-FakePnPOption ListFails $true
$failFolder = Join-Path $runFolder 'fail'
& $script -SiteUrl $Sample.Site -OutputFolder $failFolder 6>$null
$failBook = @(Get-ResultFiles $failFolder)
Assert-Equal 1 $failBook.Count 'Results saved after a failure'
Assert-Equal 'INCOMPLETE' (Get-XlsxSheets $failBook[0].FullName)['Not scanned'][1][0] 'Results say INCOMPLETE'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out after a failure'

# ---- A report that cannot be saved: the script still signs out ----
Reset-FakePnP @{}
$blocker = Join-Path $runFolder 'not-a-folder.txt'
Set-Content -LiteralPath $blocker -Value 'x'
$saveFailed = $false
try { & $script -SiteUrl $Sample.Site -OutputFolder (Join-Path $blocker 'sub') 6>$null }
catch { $saveFailed = $true }  # the save error is the expected result here
Assert-Equal $true $saveFailed 'Save error reaches the person'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out when the report could not be saved'

Remove-Item -LiteralPath $runFolder -Recurse -Force
