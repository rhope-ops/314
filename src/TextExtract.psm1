<#
TextExtract.psm1 - turns file bytes or a Teams post body into plain text, in memory.

  Purpose:     Plain text from Word, Excel and PowerPoint files (.docx/.xlsx/.pptx and macro versions),
               text files and HTML. Other types (PDF, images, legacy .doc/.xls/.ppt, scripts and
               anything else) are reported as NOT SCANNED, so a person checks them by hand.
  Inputs:      file bytes and the file name (for its extension); HTML text of a Teams post.
  Outputs:     an object with Status ('OK' or 'NOT SCANNED'), Reason and Text.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression

# =====================================================================================
# Settings: which files are read, and which parts of an Office file hold text
# =====================================================================================

# Plain text types, read as they are.
$script:PlainTextExtensions = @('.txt', '.csv', '.md', '.log', '.json', '.xml', '.tsv')

# HTML types, read and stripped of tags.
$script:HtmlExtensions = @('.html', '.htm')

# Office Open XML types (a zip of XML parts, ECMA-376) and the parts that hold their text:
# body, headers, footers, comments, notes; cell text and shared strings; slides and speaker notes.
$script:OfficeParts = @{
    '.docx' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.docm' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.xlsx' = '^xl/(sharedStrings\.xml|worksheets/sheet\d+\.xml|comments\d*\.xml)$'
    '.xlsm' = '^xl/(sharedStrings\.xml|worksheets/sheet\d+\.xml|comments\d*\.xml)$'
    '.pptx' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    '.pptm' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
}

# =====================================================================================
# Markup to text
# =====================================================================================

function Convert-XmlToText {
    <#
    .SYNOPSIS
      Strips Office XML to plain text. Paragraph, row and cell ends become line breaks or spaces;
      every other tag is removed with no space, so text that Word splits into several runs
      ("123-45-" and "6789") joins again.
    .PARAMETER Xml
      The XML text of one part.
    .OUTPUTS
      [string]
    #>
    param([string]$Xml)
    $text = $Xml -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
    $text = $text -replace '</w:p>|</a:p>|</row>', "`n"
    $text = $text -replace '</si>|</c>|</w:tc>|</a:tc>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

function Convert-HtmlToText {
    <#
    .SYNOPSIS
      Strips HTML (Teams post bodies, .html files) to plain text. Script and style blocks are dropped;
      block ends become line breaks.
    .PARAMETER Html
      The HTML text.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $text = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ' '
    $text = $text -replace '(?i)<br\s*/?>|</p>|</div>|</li>|</tr>', "`n"
    $text = $text -replace '(?i)</td>|</th>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

# =====================================================================================
# Reading bytes
# =====================================================================================

function Get-ZipPartText {
    <#
    .SYNOPSIS
      Opens an Office file (a zip) in memory and returns the text of the parts whose names match.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER PartPattern
      Regular expression on the part name (from $OfficeParts).
    .OUTPUTS
      [string] the parts' text, joined with line breaks.
    #>
    param([byte[]]$Bytes, [string]$PartPattern)
    $memory = New-Object System.IO.MemoryStream(, $Bytes)
    $zip = New-Object System.IO.Compression.ZipArchive($memory, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match $PartPattern } | Sort-Object FullName)) {
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $texts.Add((Convert-XmlToText $reader.ReadToEnd())) } finally { $reader.Dispose() }
        }
        return ($texts -join "`n")
    }
    finally { $zip.Dispose(); $memory.Dispose() }
}

function Get-PlainText {
    <#
    .SYNOPSIS
      Decodes a text file. A UTF-8 or UTF-16 byte order mark is honoured; no mark means UTF-8.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string]
    #>
    param([byte[]]$Bytes)
    $reader = New-Object System.IO.StreamReader((New-Object System.IO.MemoryStream(, $Bytes)), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Test-SupportedFile {
    # True when Get-TextFromBytes can read this file type (so other types are not downloaded).
    param([string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($script:PlainTextExtensions -contains $extension) -or ($script:HtmlExtensions -contains $extension) -or
        $script:OfficeParts.ContainsKey($extension)
}

function Get-TextFromBytes {
    <#
    .SYNOPSIS
      Returns the text of one file, or why it was not scanned.
    .PARAMETER Bytes
      The file bytes.
    .PARAMETER FileName
      The file name; only its extension is used.
    .OUTPUTS
      [pscustomobject] Status ('OK' or 'NOT SCANNED'), Reason, Text.
    #>
    param([byte[]]$Bytes, [string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    $result = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = '' }
    try {
        if ($script:PlainTextExtensions -contains $extension) { $result.Text = Get-PlainText $Bytes }
        elseif ($script:HtmlExtensions -contains $extension) { $result.Text = Convert-HtmlToText (Get-PlainText $Bytes) }
        elseif ($script:OfficeParts.ContainsKey($extension)) { $result.Text = Get-ZipPartText $Bytes $script:OfficeParts[$extension] }
        else {
            $result.Status = 'NOT SCANNED'
            $result.Reason = "File type '$extension' is not read by this script: check by hand"
        }
    }
    catch {
        # A damaged or protected file must not stop the scan: report it and go on.
        $result.Status = 'NOT SCANNED'
        $result.Reason = 'Could not read the file (damaged or protected?): ' + $_.Exception.Message
    }
    return $result
}

Export-ModuleMember -Function Get-TextFromBytes, Convert-HtmlToText, Test-SupportedFile
