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
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Test-TeamSiteUrl 'https://tenant.example.com/:w:/r/teams/TeamA/Shared%20Documents/a.docx?d=w1') `
    'A file Copy link is cut to the site'
$refused = $false
try { Test-TeamSiteUrl 'https://tenant.example.com/:w:/s/TeamA/EAbc' | Out-Null } catch { $refused = $true }  # expected refusal
Assert-Equal $true $refused 'A sharing link that does not name the site is refused'

# ---- Teams link: the group ID is read; the site is looked up, checked, and only then scanned ----
$groupId = '11111111-2222-3333-4444-555555555555'
$teamLink = 'https://teams.example.com/l/team/19%3aabc%40thread.tacv2/conversations?groupId=' + $groupId + '&tenantId=x'
$channelLink = 'https://teams.example.com/l/channel/19%3aabc%40thread.tacv2/General?groupId=' + $groupId + '&tenantId=x'
Assert-Equal $groupId (Get-ScanTarget $teamLink).GroupId 'Team link gives the group ID'
Assert-Equal $groupId (Get-ScanTarget $channelLink).GroupId 'Channel link gives the group ID'
Assert-Equal $null (Get-ScanTarget $Sample.Site).GroupId 'Site address has no group ID'
Assert-Equal $groupId (Get-ScanTarget ('https://teams.example.com/l/channel/link?groupId=' + $groupId)).GroupId 'The page form of a link is read'
$refused = $false
try { Get-ScanTarget 'https://teams.example.com/l/team/19%3aabc/conversations?tenantId=x' | Out-Null } catch { $refused = $true }  # expected
Assert-Equal $true $refused 'Teams link without groupId refused'
$realRoot = & (Get-Module ChannelScan) { $script:SharePointRoot }
& (Get-Module ChannelScan) { $script:SharePointRoot = 'https://tenant.example.com' }   # the test tenant's root
Reset-FakePnP @{}
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Lookup gives the team site'
Assert-Equal 'https://tenant.example.com https://tenant.example.com/teams/TeamA' ((Get-FakePnPConnection) -join ' ') 'Root once, then the team site'
Assert-Equal $true (Test-FakePnPCall ('Invoke-PnPSPRestMethod /_api/GroupSiteManager/GetSiteStatus ' + $groupId)) 'Lookup asks for that one group'
Reset-FakePnP @{}
Set-FakeLookup -LookupSite 'throw' -TypedSite 'https://tenant.example.com/teams/TeamA/Shared%20Documents'
Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) 'Failed lookup asks for the address'
Assert-Equal $true (Test-FakePnPCall 'Read-Host') 'Address asked after a failed lookup'
Reset-FakePnP @{}
Set-FakeLookup -LookupSite 'https://other.example.com/teams/TeamA'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: asks, gets nothing
Assert-Equal $true $refused 'Lookup on another host not used'
Assert-Equal 'https://tenant.example.com' ((Get-FakePnPConnection) -join ' ') 'No sign-in to a site on another host'
Reset-FakePnP @{}
Set-FakeLookup -LookupSite 'https://tenant.example.com/teams/TeamB' -SiteGroupId '99999999-2222-3333-4444-555555555555'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: another team's site
Assert-Equal $true $refused 'Site of another team refused'
Assert-Equal $true (Test-FakePnPCall 'Disconnect-PnPOnline') 'Signed out of the other team site'
foreach ($form in @('plain', 'json')) {
    Reset-FakePnP @{}
    Set-FakeLookup -LookupSite 'https://tenant.example.com/teams/TeamA' -Form $form
    Assert-Equal 'https://tenant.example.com/teams/TeamA' (Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null) "Lookup answer read in $form form"
}
foreach ($case in @(@{ Site = ''; Status = 0; Name = 'empty SiteUrl' }, @{ Site = 'https://tenant.example.com/teams/TeamA'; Status = 1; Name = 'site not ready' })) {
    Reset-FakePnP @{}
    Set-FakeLookup -LookupSite $case.Site -Status $case.Status -TypedSite 'https://tenant.example.com/teams/TeamA'
    $null = Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null
    Assert-Equal $true (Test-FakePnPCall 'Read-Host') "Address asked when the lookup gives $($case.Name)"
}
Reset-FakePnP @{}
Set-FakeLookup -LookupSite '' -TypedSite 'not an address'
$refused = $false
try { Connect-TeamSite (Get-ScanTarget $teamLink) 6>$null | Out-Null } catch { $refused = $true }  # expected: bad typed address
Assert-Equal $true ($refused -and (Test-FakePnPCall 'Disconnect-PnPOnline')) 'Bad typed address: signed out of the root'
Assert-Equal $true (Get-ScanTarget $channelLink).IsChannel 'Channel link marked as a channel'
Reset-FakePnP @{}
Assert-Equal $Sample.Site (Connect-TeamSite (Get-ScanTarget $Sample.Site) 6>$null) 'Site address: no lookup'
Assert-Equal $Sample.Site ((Get-FakePnPConnection) -join ' ') 'Site address: one sign-in, to that site'
& (Get-Module ChannelScan) { param($root) $script:SharePointRoot = $root } $realRoot   # back to the real root
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
$rows = New-Object System.Collections.Generic.List[object]
$null = Invoke-PostScan -SiteUrl $Sample.Site -Channels @("SSN $($Sample.Ssn)") -Rows $rows
Assert-Equal $false (Get-FakeProgressText).Contains($Sample.Ssn) 'Posts progress bar never shows the SSN'
Assert-Equal $false (($rows | ForEach-Object { $_.Channel }) -join ' ').Contains($Sample.Ssn) 'Posts rows never show the SSN'
