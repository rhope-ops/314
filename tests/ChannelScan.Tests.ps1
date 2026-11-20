<#
ChannelScan.Tests.ps1 - tests of ChannelScan.psm1 against the fake PnP cmdlets (FakePnP.ps1).

  Purpose:     Site address check; the progress meter text; file scan (hits, NOT SCANNED, masked file names, links);
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
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Test-TeamSiteUrl 'https://tenant.example.com/:w:/r/teams/TeamA/Shared%20Documents/a.docx?d=w1') `
    'A file Copy link is cut to the site'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Test-TeamSiteUrl ' tenant.example.com/teams/TeamA/Shared%20Documents') `
    'Address without https:// gets it'
$refused = $false
try { Test-TeamSiteUrl 'https://tenant.example.com/:w:/s/TeamA/EAbc' | Out-Null } catch { $refused = $true }  # expected refusal
Assert-Equal $true $refused 'A sharing link that does not name the site is refused'

# ---- Teams link: the group ID is read; the site address is asked for, checked, and only then scanned ----
$groupId = '11111111-2222-3333-4444-555555555555'
$teamLink = 'https://teams.example.com/l/team/19%3aabc%40thread.tacv2/conversations?groupId=' + $groupId + '&tenantId=x'
$channelLink = 'https://teams.example.com/l/channel/19%3aabc%40thread.tacv2/General?groupId=' + $groupId + '&tenantId=x'
Assert-Equal $groupId (Get-ScanTarget $teamLink).GroupId 'Team link gives the group ID'
Assert-Equal $groupId (Get-ScanTarget $channelLink).GroupId 'Channel link gives the group ID'
Assert-Equal $groupId (Get-ScanTarget ($channelLink -replace '^https://', '')).GroupId 'Teams link without https:// read'
Assert-Equal $null (Get-ScanTarget $Sample.Site).GroupId 'Site address has no group ID'
Assert-Equal $groupId (Get-ScanTarget ('https://teams.example.com/l/channel/link?groupId=' + $groupId)).GroupId 'The page form of a link is read'
$refused = $false
try { Get-ScanTarget 'https://teams.example.com/l/team/19%3aabc/conversations?tenantId=x' | Out-Null } catch { $refused = $true }  # expected
Assert-Equal $true $refused 'Teams link without groupId refused'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamA/Shared%20Documents'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Teams link: the typed address, cut to the site'
Assert-Equal $true (Test-FakePnPCall 'Read-Host') 'Teams link: the site address is asked for'
Assert-Equal 'https://tenant.example.com/teams/TeamA' ((Get-FakePnPConnection) -join ' ') 'Teams link: one sign-in, to the team site only (no tenant root)'
Assert-Equal $true (Test-FakePnPCall 'Quiet sign-in https://tenant.example.com/teams/TeamA') 'Sign-in warnings hidden'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'https://tenant.example.com/teams/TeamB' -SiteGroupId '99999999-2222-3333-4444-555555555555'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: another team's site
Assert-Equal $true $refused 'Site of another team refused'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out of the other team site'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'not an address'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: bad typed address
Assert-Equal $true $refused 'Bad typed address refused'
Assert-Equal 0 @(Get-FakePnPConnection).Count 'Bad typed address: no sign-in'
Assert-Equal $true (Get-ScanTarget $channelLink).IsChannel 'Channel link marked as a channel'
Assert-Equal $groupId (Get-ScanTarget ('https://teams.example.com/l/channel/19%3aabc/My Channel?groupId=' + $groupId)).GroupId `
    'Channel link with a raw space read'
Reset-FakePnP @{}
$printed = (Connect-TeamSite (Get-ScanTarget $channelLink) 6>&1 | Where-Object { $_ -is [System.Management.Automation.InformationRecord] }) -join ' '
Assert-Equal $true ($printed -match 'private or shared channel has its own site') 'Channel link prints the private-channel note'
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite 'cancel'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = ($_.Exception.Message -match 'prompt cancelled') }  # expected
Assert-Equal $true ($refused -and (Test-FakePnPCall 'Read-Host') -and (Test-FakePnPCall 'Disconnect-PnPOnline')) 'Failed address prompt: signed out'
$twoLinks = 'https://teams.example.com/l/team/A see also https://teams.example.com/l/team/B?groupId=' + $groupId
$refused = $false
try { Get-ScanTarget $twoLinks | Out-Null } catch { $refused = $true }  # expected: the first link has no groupId
Assert-Equal $true $refused 'A second link pasted after the first is not read'
$refused = $false
try { Get-ScanTarget ('https://teams.example.com/l/team/A see also teams.example.com/l/team/B?groupId=' + $groupId) | Out-Null } catch { $refused = $true }
Assert-Equal $true $refused 'A second link without https:// is not read either'
Reset-FakePnP @{}
$null = Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null
Assert-Equal $false (Test-FakePnPCall 'Disconnect-PnPOnline') 'Teams link: still signed in after success'
Reset-FakePnP @{}
Assert-Equal $Sample.Site (Connect-TeamSite (Get-ScanTarget $Sample.Site) 6>$null) 'Site address: used as given'
Assert-Equal $false (Test-FakePnPCall 'Read-Host') 'Site address: no address prompt'
Assert-Equal $false (Test-FakePnPCall 'Disconnect-PnPOnline') 'Site address: still signed in after success'
Assert-Equal $Sample.Site ((Get-FakePnPConnection) -join ' ') 'Site address: one sign-in, to that site'
$meter = @{ Total = 40; Done = 10; Read = 8; Skipped = 2; Hits = 3; HitItems = 2; Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromSeconds(30) } }
Assert-Equal '10 of 40 files done | read 8 | not scanned 2 | possible PII 3 in 2 files | 00:30 used, about 01:30 left' (Get-ScanMeterText $meter 'file') 'Meter text: counts, time used and left'
Assert-Equal 25 (Get-MeterPercent $meter) 'Meter percent'
$meter.Clock = [pscustomobject]@{ Elapsed = [TimeSpan]::FromMinutes(65) }
Assert-Equal $true ((Get-ScanMeterText $meter 'file') -match '\| 65:00 used, about 195:00 left$') 'Meter time past an hour does not wrap'
$meter.Total = 0
Assert-Equal 0 (Get-MeterPercent $meter) 'Meter percent with no total'
Assert-Equal '10 posts done | read 8 | not scanned 2 | possible PII 3 in 2 posts | 65:00 used' (Get-ScanMeterText $meter 'post') 'Meter text: no total, no time left'

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
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>$null) 'Files scanned'
Assert-Equal 4 (Get-FakeCallCount 'Progress done 2') 'Per-file bar closed after each file, also a skipped one'
Assert-Equal $true (Test-FakePnPCall 'Progress 2 under 1') 'Per-file bar sits under the main bar'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'SSN found in docx'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'Phone' -and $_.Item -eq 'notes.txt' }).Count 'Phone found in txt'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'NOT SCANNED' -and $_.Item -eq 'scan #A.pdf' }).Count 'pdf listed'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General/Roster.docx' `
    ($rows | Where-Object { $_.Item -eq 'Roster.docx' } | Select-Object -First 1).Link 'Link built'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/Ops/scan%20%23A.pdf' `
    ($rows | Where-Object { $_.Item -eq 'scan #A.pdf' }).Link 'Link escapes #'

# A plain SharePoint site: files at the top of the library are scanned too, also when there is no folder.
Reset-FakePnP @{ "$library/top.txt" = $Txt; "$library/General/notes.txt" = $Txt; "$library/Forms/AllItems.txt" = $Txt }
Set-FakeFolders @('Forms')
Assert-Equal 0 @(Get-ChannelFolderNames $lib).Count 'A site with no folders is not an error'
$rootRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 1 (Invoke-FileScan -Library $lib -Channels @() -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rootRows 6>$null) 'Top-of-library file scanned'
Assert-Equal '(top of the library)' (($rootRows | Select-Object -ExpandProperty Channel -Unique) -join ',') 'Top-of-library files named as such'
Assert-Equal 'top.txt' (($rootRows | Select-Object -ExpandProperty Item -Unique) -join ',') 'Only the top file, not the folder file'
Write-FileStep (6MB) 'Downloading' 10
Assert-Equal $true ((Get-FakeProgressText) -match 'This file Downloading - 6\.0 MB \(a large file') 'Large file: the per-file bar says it can take a minute'

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
$shown = (Invoke-FileScan -Library $lib -Channels @("SSN $($Sample.Ssn)") -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $false $shown.Contains($Sample.Ssn) 'Channel lines never show the SSN'
Assert-Equal 'SSN ***-**-####' (($rows | Select-Object -ExpandProperty Channel -Unique) -join ',') 'Channel name masked'
Assert-Equal $false ((($rows | ForEach-Object { $_.Link }) -join ' ').Contains('6789')) 'Channel link holds no SSN digits'
Assert-Equal $true ((Get-FakeProgressText) -match 'SSN \*\*\*-\*\*-####') 'Progress bar shows the masked channel name'
Assert-Equal $false (Get-FakeProgressText).Contains($Sample.Ssn) 'Progress bar never shows the SSN'
Reset-FakePnP @{
    "$library/General/Roster.docx"  = $Docx
    "$library/General/notes.txt"    = $Txt
    "$library/General/$ssnFileName" = [System.Text.Encoding]::UTF8.GetBytes('no PII inside')
    "$library/Ops/scan #A.pdf"      = $Txt
}

# ---- Download fallback: a PnP version without -AsMemoryStream ----
Remove-FakeMemoryStreamSwitch
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 3 (Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Sample.Site -MaxFileMB 25 -Rows $rows 6>$null) 'Fallback download scans'
Assert-Equal $true (Test-FakePnPCall 'Get-PnPContext') 'Fallback used the client context'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'Fallback finds the SSN'
. (Join-Path $PSScriptRoot 'FakePnP.ps1')  # put the full fake Get-PnPFile back for the next tests

# ---- A plain SharePoint site (no team): no posts, no second sign-in, no BLOCKED row ----
Reset-FakePnP @{}
Set-FakeSiteAnswer -TypedSite '' -SiteGroupId ([guid]::Empty)
$siteRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 0 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General') -Rows $siteRows 6>$null) 'Site without a team: no posts'
Assert-Equal 'NOT SCANNED' (($siteRows | Select-Object -ExpandProperty Detector) -join ',') 'Site without a team: one NOT SCANNED row'
Assert-Equal 0 @(Get-FakePnPConnection).Count 'Site without a team: no Graph sign-in'
Reset-FakePnP @{}
Set-FakePnPOption SiteFails $true
$siteRows = New-Object System.Collections.Generic.List[object]
Assert-Equal 0 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General') -Rows $siteRows 6>$null) 'Site group not readable: no posts'
Assert-Equal 'BLOCKED' (($siteRows | Select-Object -ExpandProperty Detector) -join ',') 'Site group not readable: one BLOCKED row'
Reset-FakePnP @{}

# ---- Posts refused, then allowed ----
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 0 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General') -Rows $rows 6>$null) 'Refused posts give 0'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'BLOCKED' }).Count 'BLOCKED row added'
Assert-Equal $false ($rows[0].Context -match '\d') 'Error text in a BLOCKED row is masked'
Set-FakePnPOption PostsAllowed $true
$rows = New-Object System.Collections.Generic.List[object]
Assert-Equal 2 (Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General', 'Ops') -Rows $rows 6>$null) 'Post and reply scanned'
Assert-Equal $true (Test-FakePnPCall ('Quiet sign-in ' + $Sample.Site)) 'Posts sign-in warnings hidden'
$lines = @(Invoke-PostScan -SiteUrl $Sample.Site -Channels @('General', 'Ops') -Rows (New-Object System.Collections.Generic.List[object]) 6>&1 |
    Where-Object { $_ -is [System.Management.Automation.InformationRecord] } | ForEach-Object { [string]$_ })
Assert-Equal $true (($lines -join ' ') -match '\[1/2\] General: 2 posts - read 2') 'Posts channel line counts its posts and replies'
Assert-Equal $true (($lines -join ' ') -match 'Posts done: 2 posts done \| read 2') 'Posts total counts posts and replies'
Assert-Equal 'DoD ID Personal e-mail NOT SCANNED' (($rows | ForEach-Object { $_.Detector }) -join ' ') 'Post hits; Ops has no Teams channel'
$rows = New-Object System.Collections.Generic.List[object]
$shown = (Invoke-PostScan -SiteUrl $Sample.Site -Channels @("SSN $($Sample.Ssn)") -Rows $rows 6>&1 | ForEach-Object { [string]$_ }) -join ' '
Assert-Equal $false $shown.Contains($Sample.Ssn) 'Posts channel lines never show the SSN'
Assert-Equal $false (Get-FakeProgressText).Contains($Sample.Ssn) 'Posts progress bar never shows the SSN'
Assert-Equal $false (($rows | ForEach-Object { $_.Channel }) -join ' ').Contains($Sample.Ssn) 'Posts rows never show the SSN'
