<#
TextExtract.psm1 - turns file bytes or a Teams post body into plain text, in memory.

  Purpose:     Text from .docx/.xlsx/.pptx (Office Open XML zip parts), plain text files and HTML.
               PDF, images, legacy .doc/.xls/.ppt and other types are reported as NOT SCANNED.
  Network:     none.
  Writes:      none (works on byte arrays in memory; nothing is saved to disk).
  Credentials: none.
  Runs on:     Windows PowerShell 5.1 and PowerShell 7 (.NET System.IO.Compression only).
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

$script:TextExtensions = @('.txt', '.csv', '.md', '.log', '.json', '.xml', '.tsv')
$script:HtmlExtensions = @('.html', '.htm')
$script:OfficeExtensions = @('.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm')

function Convert-XmlToText {
    <#
    .SYNOPSIS  Strips Office XML to text. Paragraph, cell and row ends become spaces or new lines;
               all other tags are removed with no space, so text split across runs joins again
               (Word often splits "123-45-" and "6789" into two runs).
    #>
    param([string]$Xml)
    $t = $Xml -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
    $t = $t -replace '</w:p>|</a:p>|</row>', "`n"
    $t = $t -replace '</si>|</c>|</w:tc>|</a:tc>', ' '
    $t = $t -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($t)
}

function Convert-HtmlToText {
    <#
    .SYNOPSIS  Strips HTML (Teams post bodies, .html files) to text. Block ends become new lines.
    #>
    param([AllowEmptyString()][string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $t = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ' '
    $t = $t -replace '(?i)<br\s*/?>|</p>|</div>|</li>|</tr>', "`n"
    $t = $t -replace '(?i)</td>|</th>', ' '
    $t = $t -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($t)
}

function Get-ZipPartText {
    <#
    .SYNOPSIS  Reads the zip entries whose names match a pattern and returns their joined text.
    .PARAMETER Bytes        The file bytes (an Office Open XML package is a zip).
    .PARAMETER NamePattern  Regex on the entry full name, for example '^word/(document|header\d*)\.xml$'.
    #>
    param([byte[]]$Bytes, [string]$NamePattern)
    $ms = New-Object System.IO.MemoryStream(, $Bytes)
    $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $parts = New-Object System.Collections.Generic.List[string]
        $entries = $zip.Entries | Where-Object { $_.FullName -match $NamePattern } | Sort-Object FullName
        foreach ($e in $entries) {
            $reader = New-Object System.IO.StreamReader($e.Open())
            try { $parts.Add((Convert-XmlToText $reader.ReadToEnd())) } finally { $reader.Dispose() }
        }
        return ($parts -join "`n")
    }
    finally { $zip.Dispose(); $ms.Dispose() }
}

function Get-TextFromBytes {
    <#
    .SYNOPSIS  Returns the text of one file, or why it was not scanned.
    .PARAMETER Bytes     File content.
    .PARAMETER FileName  Used only for its extension.
    .OUTPUTS   Object with Status ('OK' or 'NOT SCANNED'), Reason and Text.
    #>
    param([byte[]]$Bytes, [string]$FileName)
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    $result = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = '' }
    try {
        if ($script:TextExtensions -contains $ext -or $script:HtmlExtensions -contains $ext) {
            # StreamReader detects a UTF-8/UTF-16 byte order mark; no mark = UTF-8.
            $reader = New-Object System.IO.StreamReader((New-Object System.IO.MemoryStream(, $Bytes)), $true)
            try { $result.Text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($script:HtmlExtensions -contains $ext) { $result.Text = Convert-HtmlToText $result.Text }
        }
        elseif ($ext -like '.doc?') {
            $result.Text = Get-ZipPartText $Bytes '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
        }
        elseif ($ext -like '.xls?') {
            $result.Text = Get-ZipPartText $Bytes '^xl/(sharedStrings\.xml|worksheets/sheet\d+\.xml|comments\d*\.xml)$'
        }
        elseif ($ext -like '.ppt?') {
            $result.Text = Get-ZipPartText $Bytes '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
        }
        else {
            $result.Status = 'NOT SCANNED'
            $result.Reason = "File type '$ext' is not supported (PDF, images, legacy Office and others: check by hand)"
        }
    }
    catch {
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file: ' + $_.Exception.Message
    }
    return $result
}

function Test-SupportedFile {
    # True when Get-TextFromBytes can read this extension (used to skip downloads of other types).
    param([string]$FileName)
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($script:TextExtensions + $script:HtmlExtensions + $script:OfficeExtensions) -contains $ext
}

Export-ModuleMember -Function Get-TextFromBytes, Convert-HtmlToText, Test-SupportedFile
