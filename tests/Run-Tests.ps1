<#
Run-Tests.ps1 - tests for the Teams channel PII scan. No Pester needed (gov devices allow no extra modules).

  Purpose:     Detector true/false cases, masking, text extraction from generated Office files,
               and an end-to-end run against fake PnP cmdlets.
  Network:     none (PnP cmdlets are replaced by fakes defined here).
  Writes:      one CSV in a temporary folder, deleted at the end.
  Credentials: none.
  Test data:   all values are invented. PII-shaped samples are built from parts at run time
               (the PFM's decision, 2026-11-20), so this public file holds no complete SSN-, DoD ID-
               or phone-shaped string for content scanners to flag. Card numbers are published test numbers.
  Usage:       pwsh -NoProfile -File tests/Run-Tests.ps1   (exit code 0 = all passed)
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$src = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
Import-Module (Join-Path $src 'PiiDetectors.psm1') -Force
Import-Module (Join-Path $src 'TextExtract.psm1') -Force
Import-Module (Join-Path $src 'Report.psm1') -Force
Import-Module (Join-Path $src 'ChannelScan.psm1') -Force
Add-Type -AssemblyName System.IO.Compression

# ---------- Invented sample values, built from parts ----------
$Ssn = '123' + '-45-' + '6789'          # valid by SSA rules, invented
$SsnBare = '123' + '45' + '6789'
$SsnArea000 = '000' + '-12-' + '3456'
$SsnArea666 = '666' + '-12-' + '3456'
$SsnArea9xx = '912' + '-12-' + '3456'
$SsnGroup00 = '123' + '-00-' + '3456'
$SsnSerial0 = '123' + '-45-' + '0000'
$SsnMixed = '123' + '-45 ' + '6789'
$DodId = '12345' + '67890'
$Phone = '555-201' + '-0142'
$PhoneParen = '(555) 201' + '-0142'
$Passport = 'X12' + '34567'
$Email = 'pat.sample' + '@example.com'
$GovEmail = 'first.last' + '@agency.gov'
$MilEmail = 'first.last' + '@unit.af.' + 'mil'
$Site = 'https://tenant.example.com/sites/TeamA'

$script:Pass = 0
$script:Fail = New-Object System.Collections.Generic.List[string]

function Assert-Equal {
    # Records a pass, or a failure with the expected and actual values.
    param($Expected, $Actual, [string]$Name)
    if ("$Expected" -ceq "$Actual") { $script:Pass++ }
    else { $script:Fail.Add("$Name : expected [$Expected] got [$Actual]") }
}

function Get-Hits {
    # Shortcut: detector hits for a text as 'Detector/Confidence/Masked' strings.
    param([string]$Text)
    return ,@(Find-PiiInText $Text | ForEach-Object { '{0}/{1}/{2}' -f $_.Detector, $_.Confidence, $_.MaskedValue })
}

# ---------- Detectors ----------
Assert-Equal 'SSN/High/***-**-6789' (Get-Hits "Member SSN: $Ssn on file") 'SSN with keyword'
Assert-Equal 'SSN/Medium/***-**-6789' (Get-Hits "ref $Ssn end") 'SSN dashed, no keyword'
Assert-Equal 'SSN/Low/***-**-6789' (Get-Hits "value $SsnBare end") 'SSN bare digits'
Assert-Equal 0 (Get-Hits "SSN $SsnArea000").Count 'SSN area 000 rejected'
Assert-Equal 0 (Get-Hits "SSN $SsnArea666").Count 'SSN area 666 rejected'
Assert-Equal 0 (Get-Hits "SSN $SsnArea9xx").Count 'SSN area 9xx rejected'
Assert-Equal 0 (Get-Hits "SSN $SsnGroup00").Count 'SSN group 00 rejected'
Assert-Equal 0 (Get-Hits "SSN $SsnSerial0").Count 'SSN serial 0000 rejected'
Assert-Equal 0 (Get-Hits "mixed $SsnMixed").Count 'SSN mixed separators rejected'

Assert-Equal 'DoD ID/High/******7890' (Get-Hits "DoD ID: $DodId") 'DoD ID with keyword'
Assert-Equal 'DoD ID/High/******7890' (Get-Hits "EDIPI $DodId") 'EDIPI with keyword'
Assert-Equal 0 (Get-Hits "tracking $DodId shipped").Count '10 digits without keyword ignored'

Assert-Equal 'Payment card/High/************1111' (Get-Hits 'GTC card 4111 1111 1111 1111 exp') 'Card with keyword'
Assert-Equal 'Payment card/Medium/************4444' (Get-Hits 'number 5555-5555-5555-4444') 'Card no keyword'
Assert-Equal 0 (Get-Hits 'number 4111 1111 1111 1112').Count 'Card failing Luhn ignored'

Assert-Equal 'Date of birth/High/[date masked]' (Get-Hits 'DOB: 03/14/1985') 'DOB numeric'
Assert-Equal 'Date of birth/High/[date masked]' (Get-Hits 'Date of birth 14 Mar 1985') 'DOB day month year'
Assert-Equal 'Date of birth/High/[date masked]' (Get-Hits 'born March 14, 1985') 'DOB month name'
Assert-Equal 0 (Get-Hits 'Suspense 03/14/2027').Count 'Date without birth keyword ignored'

Assert-Equal 'Passport/Medium/*****567' (Get-Hits "Passport no. $Passport") 'Passport with keyword'
Assert-Equal 'Phone/Medium/***-***-0142' (Get-Hits "my cell $Phone") 'Phone personal'
Assert-Equal 'Phone/Low/***-***-0142' (Get-Hits "office $PhoneParen") 'Phone plain'
Assert-Equal 'Personal e-mail/Medium/***@example.com' (Get-Hits "write to $Email") 'Personal e-mail'
Assert-Equal 0 (Get-Hits $GovEmail).Count 'Gov e-mail ignored'
Assert-Equal 0 (Get-Hits $MilEmail).Count 'Mil e-mail ignored'

# Overlap: the SSN-shaped digits inside a longer card number must not be double counted.
Assert-Equal 1 (Get-Hits 'card 4111111111111111').Count 'Overlap gives one hit'

# Context snippet never shows the digits or the e-mail local part.
$ctx = (Find-PiiInText "Member SSN: $Ssn mail $Email")[0].Context
Assert-Equal $false ($ctx -match '\d') 'Context has no digits'
Assert-Equal $false ($ctx -match 'pat\.sample') 'Context has no e-mail local part'
Assert-Equal $true (Test-Luhn '4111111111111111') 'Luhn true'
Assert-Equal $false (Test-Luhn '4111111111111112') 'Luhn false'

# ---------- Text extraction ----------
function New-ZipBytes {
    # Builds an in-memory zip with the given entry name -> XML text pairs (a minimal Office file).
    param([hashtable]$Entries)
    $ms = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($k in $Entries.Keys) {
        $w = New-Object System.IO.StreamWriter($zip.CreateEntry($k).Open())
        $w.Write($Entries[$k]); $w.Dispose()
    }
    $zip.Dispose()
    return $ms.ToArray()
}

# Word splits text into runs; the SSN split over two runs must join again.
$ssnHead = $Ssn.Substring(0, 7)
$ssnTail = $Ssn.Substring(7)
$docx = New-ZipBytes @{ 'word/document.xml' = "<w:document><w:body><w:p><w:r><w:t>SSN $ssnHead</w:t></w:r><w:r><w:t>$ssnTail</w:t></w:r></w:p><w:p><w:r><w:t>Next &amp; line</w:t></w:r></w:p></w:body></w:document>" }
$t = Get-TextFromBytes -Bytes $docx -FileName 'Roster.docx'
Assert-Equal 'OK' $t.Status 'docx read'
Assert-Equal 'SSN/High/***-**-6789' (Get-Hits $t.Text) 'docx split runs joined'
Assert-Equal $true ($t.Text -match 'Next & line') 'docx entities decoded'

$xlsx = New-ZipBytes @{
    'xl/sharedStrings.xml'     = '<sst><si><t>DoD ID</t></si><si><t>Name</t></si></sst>'
    'xl/worksheets/sheet1.xml' = "<worksheet><sheetData><row><c t=`"s`"><v>0</v></c><c><v>$DodId</v></c></row></sheetData></worksheet>"
}
$t = Get-TextFromBytes -Bytes $xlsx -FileName 'Recall.xlsx'
Assert-Equal 'OK' $t.Status 'xlsx read'
Assert-Equal $true $t.Text.Contains($DodId) 'xlsx cell value read'

$pptx = New-ZipBytes @{ 'ppt/slides/slide1.xml' = '<p:sld><a:p><a:r><a:t>DOB: 03/14/1985</a:t></a:r></a:p></p:sld>' }
Assert-Equal 'Date of birth/High/[date masked]' (Get-Hits (Get-TextFromBytes -Bytes $pptx -FileName 'Brief.pptx').Text) 'pptx read'

$txt = [System.Text.Encoding]::UTF8.GetBytes("my cell $Phone")
Assert-Equal 'Phone/Medium/***-***-0142' (Get-Hits (Get-TextFromBytes -Bytes $txt -FileName 'notes.txt').Text) 'txt read'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $txt -FileName 'scan.pdf').Status 'pdf not scanned'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $txt -FileName 'broken.docx').Status 'corrupt docx not scanned'
$html = (Convert-HtmlToText "<div><p>SSN $Ssn</p></div>here") -replace '\s+', ' '
Assert-Equal "SSN $Ssn here" $html.Trim() 'html stripped'

# ---------- Site URL check and channel pick ----------
Assert-Equal $Site (Test-TeamSiteUrl "$Site/Shared%20Documents/Forms/AllItems.aspx") 'URL trimmed to site'
$bad = $false
try { Test-TeamSiteUrl 'http://example.com/x' | Out-Null } catch { $bad = $true }
Assert-Equal $true $bad 'Bad URL refused'
Assert-Equal 'General' (Select-ChannelFolders -Available @('General', 'Ops') -Requested 'general') 'Channel by name, any case'
Assert-Equal 'General Ops' (Select-ChannelFolders -Available @('General', 'Ops') -Requested 'All') 'Channel All'

# ---------- End to end with fake PnP cmdlets ----------
# The fakes are global functions, so the module's calls to these names reach them.
$fakeFiles = @{
    '/sites/TeamA/Shared Documents/General/Roster.docx' = $docx
    '/sites/TeamA/Shared Documents/General/notes.txt'   = $txt
    '/sites/TeamA/Shared Documents/Ops/scan.pdf'        = $txt
}
function global:Get-PnPList { param($Identity) if ($Identity -eq 'Shared Documents') { [pscustomobject]@{ Title = 'Documents' } } }
function global:Get-PnPProperty { param($ClientObject, $Property) [pscustomobject]@{ ServerRelativeUrl = '/sites/TeamA/Shared Documents' } }
function global:Get-PnPFolderItem { param($FolderSiteRelativeUrl, $ItemType) @('General', 'Ops', 'Forms') | ForEach-Object { [pscustomobject]@{ Name = $_ } } }
function global:Get-PnPListItem {
    param($List, $PageSize, $Fields)
    @{ FileLeafRef = 'General'; FileRef = '/sites/TeamA/Shared Documents/General'; FSObjType = 1; File_x0020_Size = $null; Modified = $null }
    foreach ($k in $fakeFiles.Keys) {
        @{ FileLeafRef = ($k -split '/')[-1]; FileRef = $k; FSObjType = 0; File_x0020_Size = "$($fakeFiles[$k].Length)"; Modified = [datetime]'2026-11-01' }
    }
}
function global:Get-PnPFile { param($Url, [switch]$AsMemoryStream) New-Object System.IO.MemoryStream(, $fakeFiles[$Url]) }
function global:Get-PnPSite { param($Includes) [pscustomobject]@{ GroupId = [guid]'11111111-2222-3333-4444-555555555555' } }
function global:Connect-PnPOnline { param($Url, [switch]$Interactive, $AzureEnvironment) throw 'AADSTS65001: The user or administrator has not consented (fake)' }

$lib = Get-TeamDocumentLibrary
Assert-Equal 'General Ops' (Get-ChannelFolderNames $lib) 'Folders without Forms'
$rows = New-Object System.Collections.Generic.List[object]
$n = Invoke-FileScan -Library $lib -Channels @('General', 'Ops') -SiteUrl $Site -MaxFileMB 25 -Rows $rows
Assert-Equal 2 $n 'Files scanned'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'SSN' -and $_.Item -eq 'Roster.docx' }).Count 'SSN found in docx'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'Phone' -and $_.Item -eq 'notes.txt' }).Count 'Phone found in txt'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'NOT SCANNED' -and $_.Item -eq 'scan.pdf' }).Count 'pdf listed'
Assert-Equal 'https://tenant.example.com/sites/TeamA/Shared%20Documents/General/Roster.docx' `
    ($rows | Where-Object { $_.Item -eq 'Roster.docx' } | Select-Object -First 1).Link 'Link built'

$p = Invoke-PostScan -SiteUrl $Site -Channels @('General') -Rows $rows
Assert-Equal 0 $p 'Posts blocked gives 0'
Assert-Equal 1 @($rows | Where-Object { $_.Detector -eq 'BLOCKED' }).Count 'BLOCKED row added'

# Posts allowed: fake Graph returns one post with a reply.
function global:Connect-PnPOnline { param($Url, [switch]$Interactive, $AzureEnvironment) }
function global:Get-PnPTeamsChannel { param($Team) [pscustomobject]@{ DisplayName = 'General'; Id = '19:fake' } }
function global:Get-PnPTeamsChannelMessage {
    param($Team, $Channel)
    [pscustomobject]@{ Id = '1001'; CreatedDateTime = [datetime]'2026-11-02'; Body = [pscustomobject]@{ Content = "<p>DoD ID $DodId</p>" } }
}
function global:Get-PnPTeamsChannelMessageReply {
    param($Team, $Channel, $Message)
    [pscustomobject]@{ Id = '1002'; CreatedDateTime = [datetime]'2026-11-03'; Body = [pscustomobject]@{ Content = "reply: $Email" } }
}
$rows2 = New-Object System.Collections.Generic.List[object]
Assert-Equal 2 (Invoke-PostScan -SiteUrl $Site -Channels @('General') -Rows $rows2) 'Post and reply scanned'
Assert-Equal 'DoD ID Personal e-mail' (($rows2 | ForEach-Object { $_.Detector }) -join ' ') 'Post hits'

# Report: CSV holds masked values only.
$out = Join-Path ([System.IO.Path]::GetTempPath()) ('piitest-' + [guid]::NewGuid())
$csv = Export-ScanCsv -Rows ($rows.ToArray() + $rows2.ToArray()) -OutputFolder $out
$content = Get-Content -LiteralPath $csv -Raw
Assert-Equal $false ($content.Contains($Ssn) -or $content.Contains($DodId) -or $content.Contains('pat.sample')) 'CSV has no full values'
Assert-Equal $true $content.Contains('***-**-6789') 'CSV has masked SSN'
Write-ScanSummary -Rows ($rows.ToArray() + $rows2.ToArray()) -ItemsScanned 4 -CsvPath $csv 6>$null
$emptyCsv = Export-ScanCsv -Rows @() -OutputFolder $out
Assert-Equal 1 @(Get-Content -LiteralPath $emptyCsv).Count 'Empty CSV has header'
Assert-Equal $true ($emptyCsv -ne $csv) 'Second report does not overwrite the first'

# Full script, start to end, with the fakes (posts allowed). Console output goes to the test log.
function global:Connect-PnPOnline { param($Url, [switch]$UseWebLogin, [switch]$Interactive, $AzureEnvironment) }
function global:Disconnect-PnPOnline { }
$e2e = Join-Path $out 'e2e'
& (Join-Path $src 'Find-ChannelPII.ps1') -SiteUrl $Site -Channel All -IncludePosts -OutputFolder $e2e
$report = @(Get-ChildItem -LiteralPath $e2e -Filter 'PIIScan-*.csv')
Assert-Equal 1 $report.Count 'Script wrote one CSV'
$e2eRows = @(Import-Csv -LiteralPath $report[0].FullName)
Assert-Equal 'DoD ID NOT SCANNED Personal e-mail Phone SSN' (($e2eRows | ForEach-Object { $_.Detector } | Sort-Object -Unique) -join ' ') 'Script found every planted type'
Remove-Item -LiteralPath $out -Recurse -Force

# ---------- Result ----------
foreach ($f in $script:Fail) { Write-Host "FAIL $f" -ForegroundColor Red }
Write-Host ('TESTS passed={0} failed={1}' -f $script:Pass, $script:Fail.Count)
if ($script:Fail.Count -gt 0) { exit 1 }
