<#
ChannelScan.psm1 - reads one team's channel files (and, if asked, posts) and runs the PII detectors.

  Purpose:     The SharePoint and Teams steps of Find-ChannelPII.ps1, one function per step.
  Network:     only the site URL the user gives (SharePoint) and, for posts, Microsoft Graph for that
               one team, through PnP.PowerShell 1.12.0. No other sites or teams are listed.
  Writes:      none in SharePoint or Teams (read-only calls only). Adds rows to an in-memory list.
  Credentials: the PnP connection the caller opened (own CAC sign-in); posts open a second,
               interactive sign-in. Nothing is stored.
  Runs on:     Windows PowerShell 5.1 and PowerShell 7.
  Depends on:  PiiDetectors.psm1, TextExtract.psm1, Report.psm1 (imported by the caller).
#>

Set-StrictMode -Version 2.0

# Default library of a team site. ListPipeBind finds it by title, or by URL on sites that renamed it.
$script:LibraryUrlName = 'Shared Documents'

function Test-TeamSiteUrl {
    <#
    .SYNOPSIS  Checks that the URL is one SharePoint site (https://host/sites/name or /teams/name)
               and returns it without a trailing path, so the scan can only touch that site.
    #>
    param([string]$Url)
    $m = [regex]::Match(($Url.Trim()), '^https://[A-Za-z0-9.-]+/(sites|teams)/[^/?#\s]+', 'IgnoreCase')
    if (-not $m.Success) {
        throw "Not a team site URL: '$Url'. Expected https://<tenant>.sharepoint-mil.us/sites/<team-site>."
    }
    return $m.Value
}

function Get-TeamDocumentLibrary {
    <#
    .SYNOPSIS  Finds the team's document library and returns its identity and server-relative URL.
    #>
    $list = Get-PnPList -Identity $script:LibraryUrlName
    if (-not $list) { $list = Get-PnPList -Identity 'Documents' }
    if (-not $list) { throw "No 'Shared Documents' library on this site. Is this the team's site?" }
    $rootUrl = (Get-PnPProperty -ClientObject $list -Property RootFolder).ServerRelativeUrl
    $siteRel = ($rootUrl -replace '^.*?/(sites|teams)/[^/]+/', '')
    [pscustomobject]@{ List = $list; ServerRelativeUrl = $rootUrl.TrimEnd('/'); SiteRelativeUrl = $siteRel }
}

function Get-ChannelFolderNames {
    # Top-level folders of the library = standard channels (Teams makes one folder per channel).
    param($Library)
    $items = Get-PnPFolderItem -FolderSiteRelativeUrl $Library.SiteRelativeUrl -ItemType Folder
    return @($items | ForEach-Object { $_.Name } | Where-Object { $_ -ne 'Forms' } | Sort-Object)
}

function Select-ChannelFolders {
    <#
    .SYNOPSIS  Returns the channel folders to scan: 'All', one named channel, or a pick from a list.
    .PARAMETER Available  Folder names found in the library.
    .PARAMETER Requested  -Channel value; empty = ask the user.
    #>
    param([string[]]$Available, [string]$Requested)
    if ($Available.Count -eq 0) { throw 'No channel folders found in the library.' }
    if ($Requested -eq 'All') { return $Available }
    if ($Requested) {
        $hit = @($Available | Where-Object { $_ -ieq $Requested })
        if ($hit.Count -eq 0) { throw "Channel folder '$Requested' not found. Found: $($Available -join ', ')" }
        return $hit
    }
    Write-Host 'Channel folders on this site:'
    for ($i = 0; $i -lt $Available.Count; $i++) { Write-Host ('  {0}. {1}' -f ($i + 1), $Available[$i]) }
    $answer = Read-Host 'Type a number, several numbers with commas, or A for all'
    if ($answer.Trim() -ieq 'A') { return $Available }
    $picked = foreach ($part in ($answer -split ',')) {
        $n = 0
        if ([int]::TryParse($part.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $Available.Count) { $Available[$n - 1] }
        else { throw "Not a valid choice: '$part'" }
    }
    return @($picked | Select-Object -Unique)
}

function Get-FileBytes {
    <#
    .SYNOPSIS  Downloads one file into memory (never to disk) and returns its bytes.
    .NOTES     Uses Get-PnPFile -AsMemoryStream; if this PnP version lacks that switch, falls back to
               the CSOM OpenBinaryStream call on the same connection.
    #>
    param([string]$ServerRelativeUrl)
    try {
        $stream = Get-PnPFile -Url $ServerRelativeUrl -AsMemoryStream
        return $stream.ToArray()
    }
    catch [System.Management.Automation.ParameterBindingException] {
        $ctx = Get-PnPContext
        $file = $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl)
        $data = $file.OpenBinaryStream()
        $ctx.ExecuteQuery()
        $ms = New-Object System.IO.MemoryStream
        try { $data.Value.CopyTo($ms); return $ms.ToArray() } finally { $ms.Dispose() }
    }
}

function Get-FieldValue {
    # Reads a field from a PnP list item (or a hashtable in tests) and returns it as a string.
    param($Item, [string]$Name)
    $v = $Item[$Name]
    if ($null -eq $v) { return '' }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm') }
    return [string]$v
}

function Invoke-FileScan {
    <#
    .SYNOPSIS  Scans every file in the chosen channel folders and adds report rows.
    .PARAMETER Library    Output of Get-TeamDocumentLibrary.
    .PARAMETER Channels   Channel folder names to scan.
    .PARAMETER SiteUrl    The site URL (used to build links in the report).
    .PARAMETER MaxFileMB  Larger files are reported as NOT SCANNED.
    .PARAMETER Rows       List that receives the report rows.
    .OUTPUTS   The number of files scanned (int).
    #>
    param($Library, [string[]]$Channels, [string]$SiteUrl, [int]$MaxFileMB, $Rows)
    $origin = ([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $fields = 'FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified'
    $items = @(Get-PnPListItem -List $Library.List -PageSize 500 -Fields $fields)
    $scanned = 0
    foreach ($ch in $Channels) {
        $prefix = $Library.ServerRelativeUrl + '/' + $ch + '/'
        $files = @($items | Where-Object { (Get-FieldValue $_ 'FSObjType') -eq '0' -and (Get-FieldValue $_ 'FileRef').StartsWith($prefix, 'OrdinalIgnoreCase') })
        $n = 0
        foreach ($f in $files) {
            $n++
            $name = Get-FieldValue $f 'FileLeafRef'
            $ref = Get-FieldValue $f 'FileRef'
            $item = $ref.Substring($prefix.Length)
            $link = $origin + [uri]::EscapeUriString($ref)
            $mod = Get-FieldValue $f 'Modified'
            Write-Progress -Activity "Scanning files in '$ch'" -Status $item -PercentComplete ([int](100 * $n / $files.Count))

            $size = 0L
            [void][long]::TryParse((Get-FieldValue $f 'File_x0020_Size'), [ref]$size)
            if (-not (Test-SupportedFile $name)) {
                $Rows.Add((New-ReportRow 'File' $ch $item $link $mod 'NOT SCANNED' '' '' 'File type not supported: check by hand'))
                continue
            }
            if ($size -gt ($MaxFileMB * 1MB)) {
                $Rows.Add((New-ReportRow 'File' $ch $item $link $mod 'NOT SCANNED' '' '' "Larger than $MaxFileMB MB: check by hand"))
                continue
            }
            try { $bytes = Get-FileBytes $ref }
            catch {
                $Rows.Add((New-ReportRow 'File' $ch $item $link $mod 'NOT SCANNED' '' '' ('Download failed: ' + $_.Exception.Message)))
                continue
            }
            $text = Get-TextFromBytes -Bytes $bytes -FileName $name
            $bytes = $null
            if ($text.Status -ne 'OK') {
                $Rows.Add((New-ReportRow 'File' $ch $item $link $mod 'NOT SCANNED' '' '' $text.Reason))
                continue
            }
            $scanned++
            foreach ($h in @(Find-PiiInText $text.Text)) {
                $Rows.Add((New-ReportRow 'File' $ch $item $link $mod $h.Detector $h.Confidence $h.MaskedValue $h.Context))
            }
        }
        Write-Progress -Activity "Scanning files in '$ch'" -Completed
    }
    return $scanned
}

function Add-PostRows {
    # Runs the detectors on one post or reply and adds its rows. Returns 1 (one item scanned).
    param($Message, [string]$Channel, [string]$Kind, $Rows)
    $body = ''
    if ($Message.Body) { $body = Convert-HtmlToText ([string]$Message.Body.Content) }
    $when = ''
    if ($Message.CreatedDateTime) { $when = ([datetime]$Message.CreatedDateTime).ToString('yyyy-MM-dd HH:mm') }
    $item = '{0} {1} (id {2})' -f $Kind, $when, $Message.Id
    foreach ($h in @(Find-PiiInText $body)) {
        $Rows.Add((New-ReportRow 'Post' $Channel $item '' $when $h.Detector $h.Confidence $h.MaskedValue $h.Context))
    }
    return 1
}

function Invoke-PostScan {
    <#
    .SYNOPSIS  Tries to scan the posts and replies of the chosen channels through Microsoft Graph.
    .DESCRIPTION
      Opens a second sign-in (interactive, DoD cloud) because the web login gives no Graph token.
      Reading channel messages needs the ChannelMessage.Read.All permission with admin consent.
      If any step is refused, adds one 'BLOCKED' row with the reason and returns; the file results stay.
    .OUTPUTS   The number of posts and replies scanned (int).
    #>
    param([string]$SiteUrl, [string[]]$Channels, $Rows)
    $scanned = 0
    try {
        $groupId = (Get-PnPSite -Includes GroupId).GroupId
        if (-not $groupId -or $groupId -eq [guid]::Empty) { throw 'This site is not connected to a Microsoft 365 group / team.' }
        Write-Host 'Posts: a second sign-in window opens for Microsoft Graph ...'
        Connect-PnPOnline -Url $SiteUrl -Interactive -AzureEnvironment USGovernmentDoD
        $teamChannels = @(Get-PnPTeamsChannel -Team $groupId)
    }
    catch {
        $Rows.Add((New-ReportRow 'Posts' '(all)' '(posts)' '' '' 'BLOCKED' '' '' ('Posts could not be read: ' + $_.Exception.Message +
                    ' | Needs admin consent for ChannelMessage.Read.All, or ask for a Microsoft Purview content search.')))
        return 0
    }
    foreach ($ch in $Channels) {
        $tc = @($teamChannels | Where-Object { $_.DisplayName -ieq $ch })
        if ($tc.Count -eq 0) {
            $Rows.Add((New-ReportRow 'Posts' $ch '(posts)' '' '' 'NOT SCANNED' '' '' 'No Teams channel with this folder name'))
            continue
        }
        try {
            $messages = @(Get-PnPTeamsChannelMessage -Team $groupId -Channel $tc[0].Id)
        }
        catch {
            $Rows.Add((New-ReportRow 'Posts' $ch '(posts)' '' '' 'BLOCKED' '' '' ('Posts could not be read: ' + $_.Exception.Message +
                        ' | Needs admin consent for ChannelMessage.Read.All, or ask for a Microsoft Purview content search.')))
            continue
        }
        $n = 0
        foreach ($m in $messages) {
            $n++
            Write-Progress -Activity "Scanning posts in '$ch'" -PercentComplete ([int](100 * $n / $messages.Count))
            $scanned += Add-PostRows $m $ch 'Post' $Rows
            try {
                foreach ($r in @(Get-PnPTeamsChannelMessageReply -Team $groupId -Channel $tc[0].Id -Message $m.Id)) {
                    $scanned += Add-PostRows $r $ch 'Reply' $Rows
                }
            }
            catch {
                $Rows.Add((New-ReportRow 'Posts' $ch ('Replies to post id ' + $m.Id) '' '' 'NOT SCANNED' '' '' ('Replies could not be read: ' + $_.Exception.Message)))
            }
        }
        Write-Progress -Activity "Scanning posts in '$ch'" -Completed
    }
    return $scanned
}

Export-ModuleMember -Function Test-TeamSiteUrl, Get-TeamDocumentLibrary, Get-ChannelFolderNames, Select-ChannelFolders,
Invoke-FileScan, Invoke-PostScan
