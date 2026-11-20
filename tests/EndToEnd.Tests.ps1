<#
EndToEnd.Tests.ps1 - runs Find-ChannelPII.ps1 from start to end against the fake site.

  Purpose:     The whole script: the report files are written, every planted type is found, and a failure
               part way still saves a report that says INCOMPLETE.
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

# ---- A full run with posts allowed ----
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
Set-FakePnPOption PostsAllowed $true
& $script -SiteUrl $Sample.Site -Channel All -IncludePosts -Report Both -OutputFolder $runFolder 6>$null
$csvFiles = @(Get-ChildItem -LiteralPath $runFolder -Filter 'PIIScan-*.csv')
Assert-Equal 1 $csvFiles.Count 'Script saved one CSV'
Assert-Equal 1 @(Get-ChildItem -LiteralPath $runFolder -Filter 'PIIScan-*.txt').Count 'Script saved one text report'
$types = (Import-Csv -LiteralPath $csvFiles[0].FullName | ForEach-Object { $_.Detector } | Sort-Object -Unique) -join ' '
Assert-Equal 'DoD ID NOT SCANNED Personal e-mail Phone SSN' $types 'Script found every planted type'

# ---- A run that fails part way still saves a report ----
Reset-FakePnP @{}
Set-FakePnPOption ListFails $true
$failFolder = Join-Path $runFolder 'fail'
& $script -SiteUrl $Sample.Site -Channel All -Report Csv -OutputFolder $failFolder 6>$null
$failCsv = @(Get-ChildItem -LiteralPath $failFolder -Filter 'PIIScan-*.csv')
Assert-Equal 1 $failCsv.Count 'Report saved after a failure'
Assert-Equal 'INCOMPLETE' (Import-Csv -LiteralPath $failCsv[0].FullName)[0].Detector 'Report says INCOMPLETE'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out after a failure'

# ---- A report that cannot be saved: the script still signs out ----
Reset-FakePnP @{}
$blocker = Join-Path $runFolder 'not-a-folder.txt'
Set-Content -LiteralPath $blocker -Value 'x'
$saveFailed = $false
try { & $script -SiteUrl $Sample.Site -Channel All -Report Csv -OutputFolder (Join-Path $blocker 'sub') 6>$null }
catch { $saveFailed = $true }  # the save error is the expected result here
Assert-Equal $true $saveFailed 'Save error reaches the person'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out when the report could not be saved'

Remove-Item -LiteralPath $runFolder -Recurse -Force
