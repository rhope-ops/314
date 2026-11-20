<#
Lint.ps1 - parser check and PSScriptAnalyzer (Windows PowerShell 5.1 syntax) on every .ps1/.psm1 in the repo.

  Network:     none.  Writes: none.  Credentials: none.
  Needs:       PSScriptAnalyzer (dev machine or CI only; not needed on the gov device).
  Usage:       pwsh -NoProfile -File tests/Lint.ps1   (exit code 0 = findings=0)
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module PSScriptAnalyzer
$settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
$files = Get-ChildItem -Path $root -Recurse -Include *.ps1, *.psm1
$all = foreach ($f in $files) {
    $tokens = $null; $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs)
    foreach ($x in $errs) { [pscustomobject]@{ RuleName = 'ParseError'; ScriptName = $f.Name; Line = $x.Extent.StartLineNumber; Message = $x.Message } }
    Invoke-ScriptAnalyzer -Path $f.FullName -Settings $settings
}
$all | Format-Table RuleName, ScriptName, Line, Message -AutoSize -Wrap | Out-String -Width 200
'LINT files=' + @($files).Count + ' findings=' + @($all).Count
if (@($all).Count -gt 0) { exit 1 }
