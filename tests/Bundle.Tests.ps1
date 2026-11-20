<#
Bundle.Tests.ps1 - runs the paste bundle (paste/33332Z-pii-spot-check.ps1) from start to end, as ISE would, against
the fake site.

  Purpose:     The text a person pastes works by itself: no module files, no script file of its own. The Settings
               block is filled in the way the site page fills it in (plain value replacement), then the text runs.
  Inputs:      paste/33332Z-pii-spot-check.ps1; $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1, after EndToEnd.Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      a copy of the filled-in bundle and its report files in a new temporary folder, deleted at the end.
  Credentials: none.
#>

$bundlePath = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'paste') '33332Z-pii-spot-check.ps1'
$bundleFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-bundle-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $bundleFolder | Out-Null
$library = '/sites/TeamA/Shared Documents'

# Fill in the Settings exactly as the page does: replace the empty values of the four setting lines.
$text = Get-Content -LiteralPath $bundlePath -Raw
$filled = $text.Replace("`$SiteUrl = ''", "`$SiteUrl = '$($Sample.Site)'").Replace("`$Channel = ''", "`$Channel = 'All'")
$filled = $filled.Replace('$IncludePosts = $false', '$IncludePosts = $true').Replace("`$Report = ''", "`$Report = 'Both'")
$filled = $filled.Replace("`$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PII-Scan')",
    "`$OutputFolder = '$bundleFolder'")
Assert-Equal $true ($filled -ne $text) 'Settings filled in'
Assert-Equal 0 ([regex]::Matches($filled, 'Import-Module|\$PSScriptRoot').Count) 'Bundle loads no module or file'

$runCopy = Join-Path $bundleFolder 'pasted.ps1'
Set-Content -LiteralPath $runCopy -Value $filled -Encoding UTF8
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
Set-FakePnPOption PostsAllowed $true
& $runCopy 6>$null
$csvFiles = @(Get-ChildItem -LiteralPath $bundleFolder -Filter 'PIIScan-*.csv')
Assert-Equal 1 $csvFiles.Count 'Bundle saved one CSV'
Assert-Equal 1 @(Get-ChildItem -LiteralPath $bundleFolder -Filter 'PIIScan-*.txt').Count 'Bundle saved one text report'
$types = (Import-Csv -LiteralPath $csvFiles[0].FullName | ForEach-Object { $_.Detector } | Sort-Object -Unique) -join ' '
Assert-Equal 'DoD ID NOT SCANNED Personal e-mail Phone SSN' $types 'Bundle found every planted type'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Bundle signed out'

Remove-Item -LiteralPath $bundleFolder -Recurse -Force
