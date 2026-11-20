<#
ChannelScan.Tests.ps1 - tests of ChannelScan.psm1 against the fake PnP cmdlets (FakePnP.ps1).

  Purpose:     Site address check; channel pick; file scan (hits, NOT SCANNED, masked file names, links);
               the download fallback without -AsMemoryStream; posts refused and posts allowed.
  Inputs:      $Sample, $Docx, $Txt (earlier test files); the fakes.
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- Site address and channel pick ----
Assert-Equal $Sample.Site (Test-TeamSiteUrl "$($Sample.Site)/Shared%20Documents/Forms/AllItems.aspx") 'Address cut to the site'
$refused = $false
try { Test-TeamSiteUrl 'http://example.com/x' | Out-Null } catch { $refused = $true }  # the refusal is the expected result
Assert-Equal $true $refused 'Address that is not a site refused'
Assert-Equal 'General' (Select-ChannelFolders -Available @('General', 'Ops') -Requested 'general') 'Channel by name, any case'
Assert-Equal 'General Ops' (Select-ChannelFolders -Available @('General', 'Ops') -Requested 'All') 'All channels'

# ---- File scan ----
$library = '/sites/TeamA/Shared Documents'
$ssnFileName = "Roster $($Sample.Ssn).txt"
Reset-FakePnP @{
    "$library/General/Roster.docx"  = $Docx
    "$library/General/notes.txt"    = $Txt
    "$library/General/$ssnFileName" = [System.Text.Encoding]::UTF8.GetBytes('no PII inside')
    "$library/Ops/scan #A.pdf"      = $Txt
}
$lib = Get-TeamDocumentLibrary
Assert-Equal 'General Ops' (Get-ChannelFolderNames $lib) 'Channel folders without Forms'
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows) 'Files scanned'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'SSN found in docx'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'Phone' -and $_.Item -eq 'notes.txt' }).Count 'Phone found in txt'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'NOT SCANNED' -and $_.Item -eq 'scan #A.pdf' }).Count 'pdf listed'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General/Roster.docx' `
    ($rows | Where-Object { $_.Item -eq 'Roster.docx' } | Select-Object -First 1).Link 'Link built'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/Ops/scan%20%23A.pdf' `
    ($rows | Where-Object { $_.Item -eq 'scan #A.pdf' }).Link 'Link escapes #'

# A file named after an SSN: reported as a hit, and neither the item nor the link shows the SSN.
$nameRows = @($rows | Where-Object { $_.Source -eq 'File name' })
Assert-Equal 1 $nameRows.Count 'SSN in a file name reported'
$allText = ($rows | ForEach-Object { $_.Item + ' ' + $_.Link + ' ' + $_.Context }) -join ' '
Assert-Equal $false $allText.Contains($Sample.Ssn) 'No full SSN in any item, link or snippet'
Assert-Equal $false $allText.Contains($Sample.SsnBare) 'No bare SSN in any link'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General' `
    ($nameRows[0].Link) 'Link of a file named with PII stops at its folder'

# A channel named with PII: masked in every row.
Reset-FakePnP @{ "$library/SSN $($Sample.Ssn)/a.txt" = $Txt }
$rows = New-Object System.Collections.Generic.List[object]
$null = Invoke-FileScan -Library $lib -Channels @("SSN $($Sample.Ssn)") -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows
Assert-Equal 'SSN ***-**-####' (($rows | Select-Object -ExpandProperty Channel -Unique) -join ',') 'Channel name masked'
Assert-Equal $false ((($rows | ForEach-Object { $_.Link }) -join ' ').Contains('6789')) 'Channel link holds no SSN digits'
Reset-FakePnP @{
    "$library/General/Roster.docx"  = $Docx
    "$library/General/notes.txt"    = $Txt
    "$library/General/$ssnFileName" = [System.Text.Encoding]::UTF8.GetBytes('no PII inside')
    "$library/Ops/scan #A.pdf"      = $Txt
}

# ---- Download fallback: a PnP version without -AsMemoryStream ----
Remove-FakeMemoryStreamSwitch
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows) 'Fallback download scans'
Assert-Equal $true (Test-FakePnPCall 'Get-PnPContext') 'Fallback used the client context'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'Fallback finds the SSN'
. (Join-Path $PSScriptRoot 'FakePnP.ps1')  # put the full fake Get-PnPFile back for the next tests

# ---- Posts refused, then allowed ----
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 0 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General') -Rows $rows) 'Refused posts give 0'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'BLOCKED' }).Count 'BLOCKED row added'
Assert-Equal $false ($rows[0].Context -match '\d') 'Error text in a BLOCKED row is masked'
Set-FakePnPOption PostsAllowed $true
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 2 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General', 'Ops') -Rows $rows) 'Post and reply scanned'
Assert-Equal 'DoD ID Personal e-mail NOT SCANNED' (($rows | ForEach-Object { $_.Detector }) -join ' ') 'Post hits; Ops has no Teams channel'
