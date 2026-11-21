<#
TextExtract.psm1 - turns file bytes into plain text, in memory.

  Purpose:     Plain text from Word, Excel and PowerPoint files (.docx/.xlsx/.pptx and macro versions),
               PDF (page text and form fields; the PAM, 2026-11-20), text files and HTML. Other types (images,
               legacy .doc/.xls/.ppt, scripts and anything else), encrypted files and PDFs without text are
               reported as NOT SCANNED, so a person checks them by hand.
  Inputs:      file bytes and the file name (for its extension).
  Outputs:     an object with Status ('OK' or 'NOT SCANNED'), Reason and Text.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Depends on:  PdfText.psm1 (Get-PdfText), imported first by Find-ChannelPII.ps1.
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
# body, headers, footers, comments, notes; worksheets and cell comments; slides and speaker notes.
# A worksheet is read cell by cell (Convert-SheetXmlToText), with its shared strings put back into their cells.
$script:OfficeParts = @{
    '.docx' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.docm' = '^word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$'
    '.xlsx' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.xlsm' = '^xl/(worksheets/sheet\d+|comments\d*)\.xml$'
    '.pptx' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
    '.pptm' = '^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|comments/comment\d+)\.xml$'
}

# First bytes of a zip (every .docx/.xlsx/.pptx) and of an OLE compound file. Office saves a file in OLE form when it
# is encrypted: a password, or a sensitivity label (CUI) with encryption; an old .doc/.xls/.ppt renamed is OLE too.
# The text of an encrypted file cannot be read without the Microsoft Information Protection SDK (not on the gov device).
$script:ZipMagic = '504B'
$script:OleMagic = 'D0CF11E0A1B11AE1'
$script:EncryptedReason = 'Encrypted, password-protected or labelled with encryption (or an old Office format with a new name): check by hand'

# =====================================================================================
# Markup to text
# =====================================================================================

# Between two cells of a worksheet row: a mark that no detector reads across, so values of neighbouring cells
# (123 | 45 | 6789) never join into one SSN, card or phone number (the PAM, 2026-11-20: Excel false positives).
$script:CellBreak = ' | '

function Convert-SheetXmlToText {
    <#
    .SYNOPSIS
      A worksheet as text: one line per row, cells split by $CellBreak. Only what a person sees in a cell is kept:
      text (shared or inline), whole numbers and formula results. Formula text is dropped, and so are numbers with
      a decimal part or an exponent (a time or a calculated value: 45231.416666666664 is no SSN or phone number).
    .PARAMETER Xml
      The worksheet XML.
    .PARAMETER Shared
      The shared strings of the workbook, in order (a cell of type "s" holds an index into them).
    .OUTPUTS
      [string]
    #>
    param([string]$Xml, [string[]]$Shared)
    # Tags may carry a namespace prefix (<x:row>, <x:c>): some tools write them so, and a missed tag would be a silent miss.
    $lines = foreach ($row in [regex]::Matches($Xml, '(?s)<(?:\w+:)?row\b[^>]*>(.*?)</(?:\w+:)?row>')) {
        $cells = foreach ($cell in [regex]::Matches($row.Groups[1].Value, '(?s)<(?:\w+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?c>)')) {
            $type = [regex]::Match($cell.Groups[1].Value, '\bt="(\w+)"').Groups[1].Value
            $inner = $cell.Groups[2].Value -replace '(?s)<(?:\w+:)?f\b[^>]*/>|<(?:\w+:)?f\b[^>]*>.*?</(?:\w+:)?f>', ''   # formula text: not data
            $value = [System.Net.WebUtility]::HtmlDecode([regex]::Match($inner, '(?s)<(?:\w+:)?v>(.*?)</(?:\w+:)?v>').Groups[1].Value)
            switch ($type) {
                's' { $index = 0; if ([int]::TryParse($value, [ref]$index) -and $index -lt $Shared.Count) { $Shared[$index] } }
                'inlineStr' { Convert-XmlToText $inner }
                'str' { $value }
                default { if ($type -notin @('b', 'e') -and $value -match '^-?\d+$') { $value } }   # whole numbers only
            }
        }
        (@($cells | Where-Object { $_ }) -join $script:CellBreak)
    }
    return ($lines -join "`n")
}

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
    # Word field codes are not shown on the page: PAGEREF _Toc123456789, HYPERLINK addresses, form IDs. They gave
    # false positives (a 9-digit bookmark number read as an SSN), so their text is dropped (the PAM, 2026-11-20).
    $text = $Xml -replace '(?s)<w:instrText\b[^>]*?(?:/>|>.*?</w:instrText>)', ''
    $text = $text -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
    $text = $text -replace '</w:p>|</a:p>|</row>', "`n"
    $text = $text -replace '</si>|</c>|</w:tc>|</a:tc>', ' '
    $text = $text -replace '<[^>]+>', ''
    return [System.Net.WebUtility]::HtmlDecode($text)
}

function Convert-HtmlToText {
    <#
    .SYNOPSIS
      Strips HTML (.html files) to plain text. Script and style blocks are dropped;
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

function Read-ZipEntry {
    # The text of one zip entry (UTF-8, as Office writes its XML).
    param($Entry)
    $reader = New-Object System.IO.StreamReader($Entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-SharedStrings {
    # The shared strings of a workbook in order (xl/sharedStrings.xml), each as plain text; none for other files.
    param($Zip)
    $entry = $Zip.GetEntry('xl/sharedStrings.xml')
    if (-not $entry) { return @() }
    foreach ($item in [regex]::Matches((Read-ZipEntry $entry), '(?s)<(?:\w+:)?si>(.*?)</(?:\w+:)?si>')) { Convert-XmlToText $item.Groups[1].Value }
}

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
        $shared = @(Get-SharedStrings $zip)
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -match $PartPattern } | Sort-Object FullName)) {
            $xml = Read-ZipEntry $entry
            if ($entry.FullName -match '^xl/worksheets/') { $texts.Add((Convert-SheetXmlToText $xml $shared)) }
            else { $texts.Add((Convert-XmlToText $xml)) }
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

function Get-NotZipReason {
    <#
    .SYNOPSIS
      Why an Office file's bytes are not a zip, from their first bytes; empty when they are one. Without this check
      the zip reader says only "End of central directory not found" (the PAM's DoD365 run, 2026-11-20).
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [string] the reason a person can act on, or ''.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return 'Empty file (0 bytes)' }
    $head = ($Bytes[0..([Math]::Min(7, $Bytes.Length - 1))] | ForEach-Object { $_.ToString('X2') }) -join ''
    if ($head.StartsWith($script:ZipMagic)) { return '' }
    if ($head.StartsWith($script:OleMagic)) { return $script:EncryptedReason }
    # Decoded with its byte order mark, if any (UTF-8 or UTF-16), so a web page with a mark is still seen.
    $start = (Get-PlainText ([byte[]]$Bytes[0..([Math]::Min(255, $Bytes.Length - 1))])).TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
    if ($start.StartsWith('<')) { return 'The download gave a web page, not the file: open it in SharePoint and check by hand' }
    return 'Not a valid Office file (damaged, or another type with an Office name): check by hand'
}

function Test-SupportedFile {
    # True when Get-TextFromBytes can read this file type (so other types are not downloaded).
    param([string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    return ($script:PlainTextExtensions -contains $extension) -or ($script:HtmlExtensions -contains $extension) -or
        $script:OfficeParts.ContainsKey($extension) -or $extension -eq '.pdf'
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
      [pscustomobject] Status ('OK' or 'NOT SCANNED'), Reason, Text, Note (a part of a read file that was not read).
    #>
    param([byte[]]$Bytes, [string]$FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    $result = [pscustomobject]@{ Status = 'OK'; Reason = ''; Text = ''; Note = '' }
    try {
        if ($script:PlainTextExtensions -contains $extension) { $result.Text = Get-PlainText $Bytes }
        elseif ($script:HtmlExtensions -contains $extension) { $result.Text = Convert-HtmlToText (Get-PlainText $Bytes) }
        elseif ($extension -eq '.pdf') {
            $pdf = Get-PdfText $Bytes
            if ($pdf.Reason) { $result.Status = 'NOT SCANNED'; $result.Reason = $pdf.Reason }
            else { $result.Text = $pdf.Text; $result.Note = $pdf.Note }
        }
        elseif ($script:OfficeParts.ContainsKey($extension)) {
            $notZip = Get-NotZipReason $Bytes
            if ($notZip) { $result.Status = 'NOT SCANNED'; $result.Reason = $notZip }
            else { $result.Text = Get-ZipPartText $Bytes $script:OfficeParts[$extension] }
        }
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
