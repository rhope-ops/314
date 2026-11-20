<#
ChannelScan.psm1 - reads one team's channel files (and, if asked, its posts) and turns them into report rows.

  Purpose:     The SharePoint and Teams steps of Find-ChannelPII.ps1, one function per step: check the
               site address, find the library, list the channel folders, pick channels, scan each file,
               try the posts.
  Inputs:      the site address the person gives; their PnP connection.
  Outputs:     report rows (Report.psm1 New-ReportRow) added to the list the caller passes in.
  Depends on:  PnP.PowerShell 1.12.0; PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1, Report.psm1
               (all imported first by Find-ChannelPII.ps1).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the one SharePoint site the person gives, and for posts Microsoft Graph for that one
               team. Nothing lists other sites or teams (no network reconnaissance).
  Writes:      none in SharePoint or Teams: read calls only. Rows go to an in-memory list.
  Credentials: the PnP connection the caller opened (own CAC sign-in). Posts open a second, interactive
               sign-in. Nothing is stored.
#>

Set-StrictMode -Version 2.0

# =====================================================================================
# Settings
# =====================================================================================

# A team site's file library: Teams names it "Documents" (URL "Shared Documents"); some sites renamed it.
$script:LibraryNames = @('Shared Documents', 'Documents')

# Teams makes one top-level folder per standard channel. "Forms" is SharePoint's own folder, not a channel.
$script:NotChannelFolders = @('Forms')

# Items per request when listing the library. SharePoint refuses views over 5,000 items
# (list view threshold); 500 keeps each request small.
$script:ListPageSize = 500

# Accepted site address: https://<host>/sites/<name> or /teams/<name>. Anything after the site name
# is cut off, so the scan can only touch that one site.
$script:SiteUrlPattern = '^https://[A-Za-z0-9.-]+/(sites|teams)/[^/?#\s]+'

# Microsoft cloud for the posts sign-in: DoD365 runs in the US Government DoD cloud.
$script:GraphCloud = 'USGovernmentDoD'

# What a blocked posts step needs, shown in the report.
$script:PostsBlockedAdvice = 'Needs admin consent for the Graph permission ChannelMessage.Read.All, ' +
                             'or ask your privacy or cybersecurity office for a Microsoft Purview content search.'

# =====================================================================================
# Site, library and channels
# =====================================================================================

function Test-TeamSiteUrl {
    <#
    .SYNOPSIS
      Checks that the address is one SharePoint site and returns it cut to the site name.
    .PARAMETER Url
      The address the person gave (it may go on to a library or a file).
    .OUTPUTS
      [string] https://<host>/sites/<name>
    #>
    param([string]$Url)
    $match = [regex]::Match($Url.Trim(), $script:SiteUrlPattern, 'IgnoreCase')
    if (-not $match.Success) {
        throw "Not a team site address: '$Url'. Expected https://<tenant>.sharepoint-mil.us/sites/<team-site>."
    }
    return $match.Value
}

function Get-TeamDocumentLibrary {
    <#
    .SYNOPSIS
      Finds the team's file library on the connected site.
    .OUTPUTS
      [pscustomobject] List (the PnP list), ServerRelativeUrl (/sites/<name>/Shared Documents),
      SiteRelativeUrl (Shared Documents).
    #>
    $list = $null
    foreach ($name in $script:LibraryNames) {
        # A name that is not on this site is normal (the library may be renamed): try the next one.
        $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
        if ($list) { break }
    }
    if (-not $list) { throw "No 'Shared Documents' library on this site. Is this the team's own site?" }
    $rootUrl = (Get-PnPProperty -ClientObject $list -Property RootFolder -ErrorAction Stop).ServerRelativeUrl.TrimEnd('/')
    $siteRelative = $rootUrl -replace '^.*?/(sites|teams)/[^/]+/', ''
    [pscustomobject]@{ List = $list; ServerRelativeUrl = $rootUrl; SiteRelativeUrl = $siteRelative }
}

function Get-ChannelFolderNames {
    <#
    .SYNOPSIS
      Lists the channel folders: the top-level folders of the library, without SharePoint's own.
    .PARAMETER Library
      Output of Get-TeamDocumentLibrary.
    .OUTPUTS
      [string[]] folder names, sorted.
    #>
    param($Library)
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $Library.SiteRelativeUrl -ItemType Folder -ErrorAction Stop
    return @($folders | ForEach-Object { $_.Name } | Where-Object { $script:NotChannelFolders -notcontains $_ } | Sort-Object)
}

function Select-ChannelFolders {
    <#
    .SYNOPSIS
      The channel folders to scan: all of them, one by name, or a pick from a numbered list.
    .PARAMETER Available
      The folder names (Get-ChannelFolderNames).
    .PARAMETER Requested
      The -Channel value: 'All', a folder name, or empty to ask.
    .OUTPUTS
      [string[]] the chosen folder names.
    #>
    param([string[]]$Available, [string]$Requested)
    if ($Available.Count -eq 0) { throw 'No channel folders found in the library.' }
    if ($Requested -eq 'All') { return $Available }
    if ($Requested) {
        $found = @($Available | Where-Object { $_ -ieq $Requested })
        if ($found.Count -eq 0) { throw "Channel folder '$Requested' not found. Found: $($Available -join ', ')" }
        return $found
    }
    Write-Host 'Channel folders on this site:'
    # Channel names are shown masked too (a channel named after a person's record must not show it).
    for ($i = 0; $i -lt $Available.Count; $i++) { Write-Host ('  {0}. {1}' -f ($i + 1), (Protect-PiiText $Available[$i])) }
    $answer = Read-Host 'Type a number, several numbers with commas (1,3), or A for all'
    if ($answer.Trim() -ieq 'A') { return $Available }
    $picked = foreach ($part in ($answer -split ',')) {
        $number = 0
        if (-not ([int]::TryParse($part.Trim(), [ref]$number) -and $number -ge 1 -and $number -le $Available.Count)) {
            throw "Not a valid choice: '$part'"
        }
        $Available[$number - 1]
    }
    return @($picked | Select-Object -Unique)
}

# =====================================================================================
# Files
# =====================================================================================

function Get-FieldValue {
    # Reads one field of a list item as text (dates as yyyy-MM-dd HH:mm); empty when missing.
    param($Item, [string]$Name)
    $value = $Item[$Name]
    if ($null -eq $value) { return '' }
    if ($value -is [datetime]) { return $value.ToString('yyyy-MM-dd HH:mm') }
    return [string]$value
}

function Get-SafeLink {
    <#
    .SYNOPSIS
      A link for the report that holds nothing to mask. It goes to the file when no folder or file name
      in the path would be changed by Protect-PiiText; else to the deepest folder whose name is clean
      (at least the library). Each part is escaped, so names with '#', '%' or spaces still work.
    .PARAMETER Origin
      https://<host>
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .OUTPUTS
      [string]
    #>
    param([string]$Origin, [string]$ServerRelativeUrl)
    $parts = $ServerRelativeUrl.TrimStart('/').Split('/')
    # The first three parts are the site and the library, which the person gave or Teams named.
    $keep = [Math]::Min(3, $parts.Count)
    while ($keep -lt $parts.Count -and (Protect-PiiText $parts[$keep]) -ceq $parts[$keep]) { $keep++ }
    return $Origin + '/' + (($parts[0..($keep - 1)] | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
}

function Get-FileBytes {
    <#
    .SYNOPSIS
      Downloads one file into memory (never to disk).
    .DESCRIPTION
      Tries Get-PnPFile -AsMemoryStream first. If this PnP version has no such switch, it uses the
      SharePoint client call OpenBinaryStream on the same connection.
    .PARAMETER ServerRelativeUrl
      /sites/<name>/<library>/<folder>/<file>
    .OUTPUTS
      [byte[]]
    #>
    param([string]$ServerRelativeUrl)
    try {
        $stream = Get-PnPFile -Url $ServerRelativeUrl -AsMemoryStream -ErrorAction Stop
        return $stream.ToArray()
    }
    catch [System.Management.Automation.ParameterBindingException] {
        # This PnP version has no -AsMemoryStream: use the client call instead (same connection, same rights).
        $context = Get-PnPContext
        $binary = $context.Web.GetFileByServerRelativeUrl($ServerRelativeUrl).OpenBinaryStream()
        $context.ExecuteQuery()
        $memory = New-Object System.IO.MemoryStream
        try { $binary.Value.CopyTo($memory); return $memory.ToArray() } finally { $memory.Dispose() }
    }
}

function Get-ChannelFiles {
    <#
    .SYNOPSIS
      The files (not folders) under one channel folder, at any depth.
    .PARAMETER AllItems
      Every item of the library (one Get-PnPListItem call for all channels).
    .PARAMETER Prefix
      /sites/<name>/<library>/<channel>/
    .OUTPUTS
      List items.
    #>
    param([object[]]$AllItems, [string]$Prefix)
    $AllItems | Where-Object {
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and (Get-FieldValue $_ 'FileRef').StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

function Invoke-OneFileScan {
    <#
    .SYNOPSIS
      Scans one file: its name, then its text. Adds report rows. Unsupported, too large, damaged
      or blocked files get a NOT SCANNED row instead.
    .PARAMETER File
      The list item.
    .PARAMETER Channel
      The channel folder name.
    .PARAMETER Prefix
      The channel folder's server-relative path with a trailing '/'.
    .PARAMETER Origin
      https://<host>
    .PARAMETER MaxFileMB
      The size limit.
    .PARAMETER Rows
      The list that receives the rows.
    .OUTPUTS
      [int] 1 when the file's text was scanned, else 0.
    #>
    param($File, [string]$Channel, [string]$Prefix, [string]$Origin, [int]$MaxFileMB, $Rows)
    $name = Get-FieldValue $File 'FileLeafRef'
    $ref = Get-FieldValue $File 'FileRef'
    $Channel = Protect-PiiText $Channel   # from here on, the channel name as the report shows it
    $modified = Get-FieldValue $File 'Modified'
    # The path is shown in the report: mask PII in it, and report a name that holds PII as a hit.
    $rawItem = $ref.Substring($Prefix.Length)
    $item = Protect-PiiText $rawItem
    $link = Get-SafeLink $Origin $ref
    foreach ($hit in @(Find-PiiInText $rawItem)) {
        $Rows.Add((New-ReportRow 'File name' $Channel $item $link $modified $hit.Detector $hit.Confidence $hit.MaskedValue 'In the file or folder name'))
    }
    $size = 0L
    [void][long]::TryParse((Get-FieldValue $File 'File_x0020_Size'), [ref]$size)
    $skipReason = ''
    if (-not (Test-SupportedFile $name)) { $skipReason = 'File type is not read by this script: check by hand' }
    elseif ($size -gt ($MaxFileMB * 1MB)) { $skipReason = "Larger than $MaxFileMB MB: check by hand" }
    if ($skipReason) { $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $skipReason)); return 0 }
    try { $text = Get-TextFromBytes -Bytes (Get-FileBytes $ref) -FileName $name }
    catch {
        # One file that cannot be downloaded must not stop the scan: note it and go on.
        $reason = Protect-PiiText ('Download failed: ' + $_.Exception.Message)
        $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $reason)); return 0
    }
    if ($text.Status -ne 'OK') { $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' (Protect-PiiText $text.Reason))); return 0 }
    foreach ($hit in @(Find-PiiInText $text.Text)) {
        $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
    return 1
}

function Invoke-FileScan {
    <#
    .SYNOPSIS
      Scans every file in the chosen channel folders and adds report rows.
    .PARAMETER Library
      Output of Get-TeamDocumentLibrary.
    .PARAMETER Channels
      The channel folder names to scan.
    .PARAMETER SiteUrl
      The site address (for the links in the report).
    .PARAMETER MaxFileMB
      Files larger than this are reported as NOT SCANNED.
    .PARAMETER Rows
      The list that receives the rows.
    .OUTPUTS
      [int] the number of files whose text was scanned.
    #>
    param($Library, [string[]]$Channels, [string]$SiteUrl, [int]$MaxFileMB, $Rows)
    $origin = ([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $fields = @('FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Size', 'Modified')
    $allItems = @(Get-PnPListItem -List $Library.List -PageSize $script:ListPageSize -Fields $fields -ErrorAction Stop)
    $scanned = 0
    foreach ($channel in $Channels) {
        $prefix = $Library.ServerRelativeUrl + '/' + $channel + '/'
        $files = @(Get-ChannelFiles $allItems $prefix)
        for ($i = 0; $i -lt $files.Count; $i++) {
            Write-Progress -Activity "Scanning files in '$channel'" -Status ('{0} of {1}' -f ($i + 1), $files.Count) -PercentComplete ([int](100 * ($i + 1) / $files.Count))
            $scanned += Invoke-OneFileScan $files[$i] $channel $prefix $origin $MaxFileMB $Rows
        }
        Write-Progress -Activity "Scanning files in '$channel'" -Completed
    }
    return $scanned
}

# =====================================================================================
# Posts (optional)
# =====================================================================================

function Add-PostRows {
    <#
    .SYNOPSIS
      Scans one post or reply (subject and body) and adds its report rows. The item names the post by
      date and id only, never by author (an author's name is PII).
    .PARAMETER Message
      The post or reply from PnP.
    .PARAMETER Channel
      The channel name.
    .PARAMETER Kind
      'Post' or 'Reply'.
    .PARAMETER Rows
      The list that receives the rows.
    .OUTPUTS
      [int] 1 (one item scanned).
    #>
    param($Message, [string]$Channel, [string]$Kind, $Rows)
    $text = ''
    if ($Message.PSObject.Properties['Subject'] -and $Message.Subject) { $text = [string]$Message.Subject + "`n" }
    if ($Message.Body) { $text += Convert-HtmlToText ([string]$Message.Body.Content) }
    $when = ''
    if ($Message.CreatedDateTime) { $when = ([datetime]$Message.CreatedDateTime).ToString('yyyy-MM-dd HH:mm') }
    $item = '{0} {1} (id {2})' -f $Kind, $when, $Message.Id
    foreach ($hit in @(Find-PiiInText $text)) {
        $Rows.Add((New-ReportRow 'Post' $Channel $item '' $when $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
    return 1
}

function Connect-TeamGraph {
    <#
    .SYNOPSIS
      Opens the second sign-in for Microsoft Graph and returns the team's id and channels.
      Throws when the site has no team or the tenant refuses the sign-in or the call.
    .PARAMETER SiteUrl
      The team's site address.
    .OUTPUTS
      [pscustomobject] GroupId, Channels.
    #>
    param([string]$SiteUrl)
    $groupId = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    if (-not $groupId -or $groupId -eq [guid]::Empty) { throw 'This site is not connected to a team.' }
    Write-Host 'Posts: a second sign-in window opens for Microsoft Graph ...'
    Connect-PnPOnline -Url $SiteUrl -Interactive -AzureEnvironment $script:GraphCloud -ErrorAction Stop
    $channels = @(Get-PnPTeamsChannel -Team $groupId -ErrorAction Stop)
    [pscustomobject]@{ GroupId = $groupId; Channels = $channels }
}

function Invoke-ChannelPostScan {
    <#
    .SYNOPSIS
      Scans the posts and replies of one channel. A refused call becomes a BLOCKED or NOT SCANNED row.
    .PARAMETER Team
      Output of Connect-TeamGraph.
    .PARAMETER Channel
      The channel folder name (Teams uses the same name for the channel).
    .PARAMETER Rows
      The list that receives the rows.
    .OUTPUTS
      [int] the number of posts and replies scanned.
    #>
    param($Team, [string]$Channel, $Rows)
    $teamsChannel = @($Team.Channels | Where-Object { $_.DisplayName -ieq $Channel })
    $Channel = Protect-PiiText $Channel   # from here on, the channel name as the report shows it
    if ($teamsChannel.Count -eq 0) { $Rows.Add((New-ReportRow 'Posts' $Channel '(posts)' '' '' 'NOT SCANNED' '' '' 'No Teams channel with this folder name')); return 0 }
    $channelId = $teamsChannel[0].Id
    try { $messages = @(Get-PnPTeamsChannelMessage -Team $Team.GroupId -Channel $channelId -ErrorAction Stop) }
    catch {
        # Posts are optional: a refusal is reported, and the file results stay.
        $Rows.Add((New-ReportRow 'Posts' $Channel '(posts)' '' '' 'BLOCKED' '' '' ((Protect-PiiText ('Posts could not be read: ' + $_.Exception.Message)) + ' | ' + $script:PostsBlockedAdvice)))
        return 0
    }
    $scanned = 0
    foreach ($message in $messages) {
        $scanned += Add-PostRows $message $Channel 'Post' $Rows
        try {
            foreach ($reply in @(Get-PnPTeamsChannelMessageReply -Team $Team.GroupId -Channel $channelId -Message $message.Id -ErrorAction Stop)) {
                $scanned += Add-PostRows $reply $Channel 'Reply' $Rows
            }
        }
        catch {
            # Replies of one post could not be read: note it and go on with the next post.
            $Rows.Add((New-ReportRow 'Posts' $Channel ('Replies to post id ' + $message.Id) '' '' 'NOT SCANNED' '' '' (Protect-PiiText ('Replies could not be read: ' + $_.Exception.Message))))
        }
    }
    return $scanned
}

function Invoke-PostScan {
    <#
    .SYNOPSIS
      Tries to scan the posts and replies of the chosen channels through Microsoft Graph. Reading posts
      needs a permission an admin may not have granted; then one BLOCKED row says so, and the file
      results stay.
    .PARAMETER SiteUrl
      The team's site address.
    .PARAMETER Channels
      The channel names to scan.
    .PARAMETER Rows
      The list that receives the rows.
    .OUTPUTS
      [int] the number of posts and replies scanned.
    #>
    param([string]$SiteUrl, [string[]]$Channels, $Rows)
    try { $team = Connect-TeamGraph $SiteUrl }
    catch {
        # Posts are optional: a refused sign-in is reported, and the file results stay.
        $Rows.Add((New-ReportRow 'Posts' '(all)' '(posts)' '' '' 'BLOCKED' '' '' ((Protect-PiiText ('Posts could not be read: ' + $_.Exception.Message)) + ' | ' + $script:PostsBlockedAdvice)))
        return 0
    }
    $scanned = 0
    foreach ($channel in $Channels) {
        Write-Progress -Activity 'Scanning posts' -Status $channel
        $scanned += Invoke-ChannelPostScan $team $channel $Rows
    }
    Write-Progress -Activity 'Scanning posts' -Completed
    return $scanned
}

Export-ModuleMember -Function Test-TeamSiteUrl, Get-TeamDocumentLibrary, Get-ChannelFolderNames, Select-ChannelFolders,
Invoke-FileScan, Invoke-PostScan
