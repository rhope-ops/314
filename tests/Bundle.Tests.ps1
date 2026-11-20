<#
Bundle.Tests.ps1 - runs the paste bundle (paste/33332Z-pii-spot-check.ps1) from start to end, as ISE would, against
the fake site.

  Purpose:     The text a person pastes works by itself: no module files, no script file of its own (the modules of
               the earlier test files are unloaded first). The Settings block is filled in the way the site page fills
               it in (plain value replacement); the text then runs dot-sourced, as ISE runs an untitled pane, and must
               leave the session's error setting and strict mode as they were.
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

# Unload the modules the earlier test files imported, so the bundle can only use its own copies of the functions.
Get-Module PiiDetectors, PiiMasking, TextExtract, Report, ChannelScan | Remove-Module -Force
Assert-Equal $false ([bool](Get-Command Find-PiiInText -ErrorAction SilentlyContinue)) 'Modules unloaded before the bundle runs'

$runCopy = Join-Path $bundleFolder 'pasted.ps1'
Set-Content -LiteralPath $runCopy -Value $filled -Encoding UTF8
Reset-FakePnP @{ "$library/General/Roster.docx" = $Docx; "$library/General/notes.txt" = $Txt; "$library/Ops/scan.pdf" = $Txt }
Set-FakePnPOption PostsAllowed $true
# ISE runs an untitled script pane in the session itself, as dot-sourcing does: run it that way, from a session with
# the default error setting and no strict mode, and check that the bundle puts both back.
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
. $runCopy 6>$null
$sessionErrorAction = $ErrorActionPreference
$strictStillOn = $false
try { $null = $notDefinedAnywhere } catch { $strictStillOn = $true }  # strict mode makes an unset variable an error
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'Continue' $sessionErrorAction 'Bundle put the session error setting back'
Assert-Equal $false $strictStillOn 'Bundle turned strict mode off at the end'
$csvFiles = @(Get-ChildItem -LiteralPath $bundleFolder -Filter 'PIIScan-*.csv')
Assert-Equal 1 $csvFiles.Count 'Bundle saved one CSV'
Assert-Equal 1 @(Get-ChildItem -LiteralPath $bundleFolder -Filter 'PIIScan-*.txt').Count 'Bundle saved one text report'
$types = (Import-Csv -LiteralPath $csvFiles[0].FullName | ForEach-Object { $_.Detector } | Sort-Object -Unique) -join ' '
Assert-Equal 'DoD ID NOT SCANNED Personal e-mail Phone SSN' $types 'Bundle found every planted type'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Bundle signed out'

# The clean-up also runs when the scan fails part way: the session's error setting comes back after the error.
Reset-FakePnP @{}
Set-FakePnPOption ListFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
. $runCopy 6>$null
$errorPathAction = $ErrorActionPreference
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'Continue' $errorPathAction 'Bundle put the session error setting back after a failed scan'
Assert-Equal $true (@(Get-ChildItem -LiteralPath $bundleFolder -Filter 'PIIScan-*.csv').Count -ge 2) 'Failed scan still saved its report'

# An error before the scan's own try (the sign-in fails) leaves the bundle through its outer 'finally' only: strict
# mode off, the session's error setting back, signed out, and the error reaches the person.
Reset-FakePnP @{}
Set-FakePnPOption ConnectFails $true
$ErrorActionPreference = 'Continue'
Set-StrictMode -Off
$outerError = ''
try { . $runCopy 6>$null } catch { $outerError = $_.Exception.Message }
$outerAction = $ErrorActionPreference
$outerStrict = $false
try { $null = $notDefinedAnywhereEither } catch { $outerStrict = $true }  # strict mode makes an unset variable an error
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Assert-Equal 'FAKE: sign-in refused' $outerError 'Sign-in error reached the person'
Assert-Equal 'Continue' $outerAction 'Outer clean-up put the session error setting back'
Assert-Equal $false $outerStrict 'Outer clean-up turned strict mode off'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Outer error path signed out'

Remove-Item -LiteralPath $bundleFolder -Recurse -Force
