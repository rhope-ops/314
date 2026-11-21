<#
PdfText.psm1 - the text of a PDF file, in memory, with .NET only (no module; the PAM, 2026-11-20).

  Purpose:     Page text (Tj/TJ text in Flate, ASCII85 or hex streams, with ToUnicode font maps), filled-in form-field
               values, comments and file info, so the detectors can check a PDF (ISO 32000). An encrypted PDF,
               a PDF with no text (a scanned image) or a file that is not a PDF gets a reason instead.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text and Reason (Get-PdfText).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters, ToUnicode maps) and OcrText.psm1 (scanned pages), imported first
               by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:PdfLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary). A literal holds at most
# 10,000 parts, so an unclosed bracket fails after a bounded scan. A "(" that no string token takes means a string
# was not read: the PDF is then NOT SCANNED, never read in part without a word (Assert-PdfGap).
$script:PdfStringPattern = '\((?>[^()\\]+|\\.|\((?<d>)|\)(?<-d>)){0,10000}(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = $script:PdfStringPattern + '|%[^\r\n]*|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'

# A TJ gap wider than this (thousandths of a text unit) is a space between words.
$script:PdfWordGap = 250

# An object stream (PDF 1.5) is unpacked only within these limits: its header of object numbers and offsets, and its
# object count (real files hold a few hundred objects per stream). More gives a "check by hand" note.
$script:MaxPackedHeader = 1MB
$script:MaxPackedObjects = 100000

# Dictionary keys whose string values are text a person wrote (Get-PdfDictText) -> the label put before the value:
# a form field's value (labelled with the field's name), a comment or note and its rich text, the file's info.
$script:PdfTextKeys = [ordered]@{ 'V' = ''; 'Contents' = 'Comment'; 'RC' = 'Comment'; 'Author' = 'Author'; 'Title' = 'Title'
    'Subject' = 'Subject'; 'Keywords' = 'Keywords' }

# The end of an inline image's bytes (BI <keys> ID <bytes> EI): white space, EI, then a delimiter or the end. After a
# BI operator (a token, never the word BI in a string) the reader steps over the bytes to here, so image bytes never
# join with page text into one token (fresh reviews, 2026-11-21).
$script:InlineImageEnd = '\sEI(?=[\s/\[<(%]|$)'

$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'

# A scanned PDF (no text layer): its JPEG page images are read with Windows OCR, at most this many (OcrText.psm1).
$script:MaxScanPages = 30
$script:PdfScanImages = @()
$script:PdfSlowReason = 'This PDF is too large or damaged to read in time: check by hand'
$script:PdfPartReason = 'Part of this PDF could not be read (a damaged or unusual string): check by hand'
$script:PdfStringError = 'PDF string not read'
$script:PdfEncryptedReason = 'Encrypted PDF (password or label): check by hand'

function Get-PdfObjects {
    <#
    .SYNOPSIS
      The objects of a PDF by number: each its dictionary text and its stream (inflated when FlateDecode; $null for
      other filters, such as images). Objects packed in object streams (PDF 1.5) are unpacked too.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary] number -> @{ Dict; Stream }, in file order.
    #>
    param([string]$Text)
    $objects = New-Object System.Collections.Specialized.OrderedDictionary
    # Each object runs from its header to the first "endobj" after it, or to the next header when "endobj" is missing.
    # Two linear passes, so a damaged file without "endobj" cannot make the search quadratic.
    $heads = @((Get-PdfRegex '(\d+)\s+\d+\s+obj\b').Matches($Text))
    $ends = @((Get-PdfRegex '\bendobj\b').Matches($Text) | ForEach-Object { $_.Index })
    $e = 0
    $split = Get-PdfRegex '(?s)^(.*?)\bstream(?:\r\n|\r|\n)(.*?)(?:\r\n|\r|\n)?endstream'
    for ($h = 0; $h -lt $heads.Count; $h++) {
        Assert-PdfTime
        $start = $heads[$h].Index + $heads[$h].Length
        $stop = $Text.Length; if ($h + 1 -lt $heads.Count) { $stop = $heads[$h + 1].Index }
        while ($e -lt $ends.Count -and $ends[$e] -lt $start) { $e++ }
        if ($e -lt $ends.Count -and $ends[$e] -lt $stop) { $stop = $ends[$e] }
        $body = $Text.Substring($start, $stop - $start)
        $stream = $null
        $parts = $split.Match($body)
        $image = $null
        if ($parts.Success) {
            $body = $parts.Groups[1].Value
            $stream = Read-PdfStream $body $parts.Groups[2].Value
            # A JPEG page image (a scanned page): kept as it is, for OCR when the PDF has no text (Get-PdfScanText).
            if ($body -match '/Subtype\s*/Image\b' -and $body -match '/Filter\s*\[?\s*/DCTDecode\s*\]?') { $image = $parts.Groups[2].Value }
        }
        elseif ($body -match '\bstream(?:\r\n|\r|\n)') { Add-PdfNote 'Part of this PDF is damaged: check by hand' }   # no endstream: cut short
        $objects[$heads[$h].Groups[1].Value] = @{ Dict = $body; Stream = $stream; Image = $image }
    }
    foreach ($packed in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/ObjStm' -and $_.Stream })) { Add-PackedObjects $objects $packed }
    return $objects
}

function Add-PackedObjects {
    <#
    .SYNOPSIS
      Adds the objects of one object stream (/Type /ObjStm: N pairs "number offset", then the objects from /First).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Packed
      One object stream of Get-PdfObjects.
    #>
    param($Objects, $Packed)
    $count = [int][regex]::Match($Packed.Dict, '/N\s+(\d+)').Groups[1].Value
    $first = [int][regex]::Match($Packed.Dict, '/First\s+(\d+)').Groups[1].Value
    if ($first -le 0 -or $first -gt $Packed.Stream.Length) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return }
    if ($first -gt $script:MaxPackedHeader -or $count -gt $script:MaxPackedObjects) { Add-PdfNote 'Part of this PDF is too large to unpack: check by hand'; return }
    $pairs = @($Packed.Stream.Substring(0, $first).Trim() -split '\s+')
    $number = 0
    for ($k = 0; $k -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        if (-not [int]::TryParse($pairs[$k], [ref]$number)) { Add-PdfNote 'Part of this PDF is damaged: check by hand'; return }
    }
    for ($k = 0; $k -lt $count -and 2 * $k + 1 -lt $pairs.Count; $k++) {
        if (($k % 2000) -eq 0) { Assert-PdfTime }
        $start = $first + [int]$pairs[2 * $k + 1]
        $end = $Packed.Stream.Length
        if (2 * $k + 3 -lt $pairs.Count) { $end = $first + [int]$pairs[2 * $k + 3] }
        if ($start -ge $end -or $end -gt $Packed.Stream.Length) { continue }
        if (-not $Objects.Contains($pairs[2 * $k])) { $Objects[$pairs[2 * $k]] = @{ Dict = $Packed.Stream.Substring($start, $end - $start); Stream = $null; Image = $null } }
    }
}

function ConvertFrom-PdfString {
    <#
    .SYNOPSIS
      The bytes of one PDF string token, as Latin-1 characters: (literal) with its escapes undone, or <hex>.
    .PARAMETER Token
      One PDF string token.
    #>
    param([string]$Token)
    if ($Token.StartsWith('<')) {
        $hex = $Token.Substring(1, $Token.Length - 2) -replace '\s', ''
        if ($hex.Length % 2) { $hex += '0' }
        return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) {
            if (($i % 200000) -eq 0) { Assert-PdfTime }
            [char][Convert]::ToInt32($hex.Substring($i, 2), 16)
        })
    }
    $undo = [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $code = $m.Value.Substring(1)
        if ($code -match '^[0-7]+$') { return [string][char]([Convert]::ToInt32($code, 8) -band 0xFF) }
        switch -CaseSensitive ($code) { 'n' { "`n" } 'r' { "`r" } 't' { "`t" } 'b' { ' ' } 'f' { ' ' } default { $code.Trim("`r", "`n") } }
    }
    return [regex]::Replace($Token.Substring(1, $Token.Length - 2), '\\(?:[0-7]{1,3}|\r?\n|.)', $undo)
}

function ConvertFrom-PdfTextString {
    # A PDF text string (a form field's name or value): UTF-16 when it starts with the FE FF mark, else Latin-1.
    param([string]$Token)
    $raw = ConvertFrom-PdfString $Token
    if ($raw.Length -ge 2 -and [int]$raw[0] -eq 0xFE -and [int]$raw[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($script:PdfLatin1.GetBytes($raw.Substring(2)))
    }
    return $raw
}

function Get-PdfFonts {
    <#
    .SYNOPSIS
      The fonts the pages name (/F1 ...) -> their ToUnicode map, or $null (simple fonts: bytes read as Latin-1).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param($Objects)
    $fonts = @{}
    foreach ($obj in $Objects.Values) {
        Assert-PdfTime
        $one = Get-PdfFontTable $obj.Dict $Objects
        foreach ($name in $one.Keys) { if (-not $fonts.ContainsKey($name)) { $fonts[$name] = $one[$name] } }
    }
    return $fonts
}

function Get-PdfFontTable {
    <#
    .SYNOPSIS
      The fonts that one dictionary's /Font resource names (/F1 -> its ToUnicode map, or $null). A page uses its own
      table, so two pages that both name a font /F1 each get theirs; ToUnicode maps are read once per font.
    .PARAMETER Dict
      A page or resources dictionary.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [hashtable]
    #>
    param([string]$Dict, $Objects)
    $fonts = @{}
    foreach ($resource in [regex]::Matches($Dict, '/Font\s*(?:<<(?<inline>[^<>]*)>>|(?<ref>\d+)\s+\d+\s+R)')) {
        $names = $resource.Groups['inline'].Value
        if ($resource.Groups['ref'].Success -and $Objects.Contains($resource.Groups['ref'].Value)) { $names = $Objects[$resource.Groups['ref'].Value].Dict }
        foreach ($font in [regex]::Matches($names, '/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R')) {
            $id = $font.Groups[2].Value
            if (-not $Objects.Contains($id)) { continue }
            $entry = $Objects[$id]
            if (-not $entry.Contains('Map')) {
                $entry['Map'] = $null
                $unicode = [regex]::Match($entry.Dict, '/ToUnicode\s+(\d+)\s+\d+\s+R')
                if ($unicode.Success -and $Objects.Contains($unicode.Groups[1].Value) -and $Objects[$unicode.Groups[1].Value].Stream) {
                    $cmap = $Objects[$unicode.Groups[1].Value]   # read once, also when many fonts share it
                    if (-not $cmap.Contains('CMap')) { $cmap['CMap'] = Get-PdfCMap $cmap.Stream }
                    $entry['Map'] = $cmap['CMap']
                }
            }
            $fonts[$font.Groups[1].Value] = $entry['Map']
        }
    }
    return $fonts
}

function ConvertFrom-PdfGlyphs {
    <#
    .SYNOPSIS
      The text of one shown string: each glyph code looked up in the font's ToUnicode map; Latin-1 without a map.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    .PARAMETER Font
      A ToUnicode map of Get-PdfFonts, or $null.
    #>
    param([string]$Raw, $Font)
    if (-not $Font) { return $Raw }
    $out = New-Object System.Text.StringBuilder
    $step = [Math]::Max(1, [int]$Font.Bytes)
    for ($i = 0; $i + $step -le $Raw.Length; $i += $step) {
        if ((($i / $step) % 1000) -eq 999) { Assert-PdfTime }
        [long]$number = 0
        for ($j = 0; $j -lt $step; $j++) { $number = $number * 256 + [int]$Raw[$i + $j] }
        $text = $Font.Map[$number.ToString('X' + (2 * $step))]
        if ($null -eq $text) {
            foreach ($range in $Font.Ranges) {
                if ($number -ge $range.Low -and $number -le $range.High) {
                    $text = ConvertFrom-Utf16Hex ($range.Prefix + (($range.Start + ($number - $range.Low)) -band 0xFFFF).ToString('X4'))
                    break
                }
            }
        }
        if ($null -ne $text) { [void]$out.Append($text) } elseif ($step -eq 1) { [void]$out.Append($Raw[$i]) }
    }
    return $out.ToString()
}

function Convert-PdfContentToText {
    <#
    .SYNOPSIS
      The text a page content stream shows: Tj, TJ, ' and " strings in the current font; a line break at each
      line move and text block end; a space at a wide TJ gap.
    .PARAMETER Content
      The (inflated) content stream.
    .PARAMETER Fonts
      Output of Get-PdfFonts or Get-PdfFontTable.
    .OUTPUTS
      [string]
    #>
    param([string]$Content, [hashtable]$Fonts)
    $out = New-Object System.Text.StringBuilder
    $stack = New-Object System.Collections.Generic.List[object]
    $array = $null
    $font = $null
    $count = 0
    $next = 0           # where the last token ended: text between tokens must hold no "(" (Assert-PdfGap)
    $inImage = $false   # after a BI operator: the next ID starts the bytes of an inline image
    $tokens = Get-PdfRegex $script:PdfTokenPattern
    $imageEnd = Get-PdfRegex $script:InlineImageEnd
    # Token by token from a position, so the bytes of an inline image can be stepped over (they are not tokens).
    for ($token = $tokens.Match($Content, 0); $token.Success; $token = $tokens.Match($Content, $next)) {
        if ((++$count % 100) -eq 0) { Assert-PdfTime }
        $t = $token.Value
        Assert-PdfGap $Content $next $token.Index
        $next = $token.Index + $token.Length
        if ($t[0] -eq '%') { continue }   # a comment
        if ($t -ceq 'BI') { $inImage = $true; $stack.Clear(); continue }
        if ($inImage) {
            if ($t -ceq 'ID') { $next = Skip-PdfInlineImage $Content $next $imageEnd; $inImage = $false }   # else: its keys
            continue
        }
        if ($t -eq '[') { $array = New-Object System.Collections.Generic.List[string]; continue }
        if ($t -eq ']') { if ($null -ne $array) { $stack.Add($array) }; $array = $null; continue }
        if ($t[0] -eq '(' -or $t[0] -eq '<' -or $t[0] -eq '/' -or $t -match '^[-+.\d]') {
            if ($null -ne $array) { $array.Add($t) } else { $stack.Add($t) }
            continue
        }
        $last = $null; if ($stack.Count -gt 0) { $last = $stack[$stack.Count - 1] }
        $shown = ''   # the string a Tj, ' or " shows: only a string token counts
        if ($last -is [string] -and $last.Length -ge 2 -and ($last[0] -eq '(' -or $last[0] -eq '<')) { $shown = ConvertFrom-PdfGlyphs (ConvertFrom-PdfString $last) $font }
        switch -CaseSensitive ($t) {
            'Tf' { if ($stack.Count -ge 2) { $font = $Fonts[([string]$stack[$stack.Count - 2]).TrimStart('/')] } }
            'Tj' { [void]$out.Append($shown) }
            "'" { [void]$out.Append("`n" + $shown) }
            '"' { [void]$out.Append("`n" + $shown) }
            'TJ' { if ($last -is [System.Collections.Generic.List[string]]) { [void]$out.Append((Convert-PdfArrayToText $last $font)) } }
            'Td' { if ($stack.Count -ge 2 -and (ConvertTo-PdfNumber $last) -ne 0) { [void]$out.Append("`n") } else { [void]$out.Append(' ') } }
            'TD' { [void]$out.Append("`n") }
            'T*' { [void]$out.Append("`n") }
            'Tm' { [void]$out.Append("`n") }
            'ET' { [void]$out.Append("`n") }
        }
        $stack.Clear()
    }
    Assert-PdfGap $Content $next $Content.Length
    return $out.ToString()
}

function Skip-PdfInlineImage {
    <#
    .SYNOPSIS
      Where the reader goes on after an inline image's ID: past the image bytes, white space and EI. Without an EI the
      rest of the stream is image bytes: a note, and the end of the stream.
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the ID operator ended.
    .PARAMETER ImageEnd
      The timed regex of $InlineImageEnd.
    .OUTPUTS
      [int] the position after EI.
    #>
    param([string]$Content, [int]$From, $ImageEnd)
    $end = $ImageEnd.Match($Content, [Math]::Min($Content.Length, $From + 1))   # ID is followed by one white-space byte
    if ($end.Success) { return ($end.Index + $end.Length) }
    Add-PdfNote 'Part of this PDF is damaged: check by hand'
    return $Content.Length
}

function Assert-PdfGap {
    <#
    .SYNOPSIS
      Throws when the text between two tokens holds a "(": a string there was not read (too long, or damaged), so
      the PDF is reported NOT SCANNED rather than read in part without a word (fresh review, 2026-11-20).
    .PARAMETER Content
      The content stream.
    .PARAMETER From
      Where the gap starts.
    .PARAMETER To
      Where the gap ends.
    #>
    param([string]$Content, [int]$From, [int]$To)
    if ($To -gt $From -and $Content.IndexOf('(', $From, $To - $From) -ge 0) { throw (New-Object System.FormatException $script:PdfStringError) }
}

function ConvertTo-PdfNumber {
    # A PDF number token as a number; 0 when it is not one.
    param($Token)
    $number = 0.0
    if (-not [double]::TryParse([string]$Token, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return 0.0 }
    return $number
}

function Convert-PdfArrayToText {
    # The text of a TJ array: its strings in the font, a space where a gap is wider than $PdfWordGap.
    param($Items, $Font)
    $out = New-Object System.Text.StringBuilder
    foreach ($item in $Items) {
        if ($item.Length -ge 2 -and ($item[0] -eq '(' -or $item[0] -eq '<')) { [void]$out.Append((ConvertFrom-PdfGlyphs (ConvertFrom-PdfString $item) $Font)) }
        elseif ((ConvertTo-PdfNumber $item) -lt -$script:PdfWordGap) { [void]$out.Append(' ') }
    }
    return $out.ToString()
}

function Get-PdfDictText {
    <#
    .SYNOPSIS
      Text kept in dictionaries, not on a page: form-field values ("name: value", so labels stay with values),
      comments and sticky notes (/Contents of an annotation, /RC rich text) and the file's info (author, title,
      subject, keywords). A value may sit in an object of its own (/V 6 0 R). A literal value that is not read throws
      (Get-PdfText: "part of this PDF could not be read"), never a silent miss (fresh reviews, 2026-11-21).
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] one line per value.
    #>
    param($Objects)
    $nameRegex = Get-PdfRegex ('/T\s*(' + $script:PdfStringPattern + ')')
    $checks = foreach ($key in $script:PdfTextKeys.Keys) {
        @{ Key = $key; Label = $script:PdfTextKeys[$key]; Value = (Get-PdfRegex ('/' + $key + '\s*(' + $script:PdfStringPattern + ')'))
           Ref = (Get-PdfRegex ('/' + $key + '\s+(\d+)\s+\d+\s+R')); Open = ('/' + $key + '\s*\(') }
    }
    foreach ($obj in $Objects.Values) {
        Assert-PdfTime
        foreach ($check in $checks) {
            if ($obj.Dict.IndexOf('/' + $check.Key) -lt 0) { continue }
            $token = ''
            $value = $check.Value.Match($obj.Dict)
            if ($value.Success) { $token = $value.Groups[1].Value }
            elseif ($obj.Dict -match $check.Open) { throw (New-Object System.FormatException $script:PdfStringError) }   # not read
            else {
                # A value in an object of its own: that object is just a string (a page's /Contents points to a stream).
                $ref = $check.Ref.Match($obj.Dict)
                if ($ref.Success -and $Objects.Contains($ref.Groups[1].Value)) {
                    $target = $Objects[$ref.Groups[1].Value].Dict.Trim()
                    if ($target.StartsWith('(') -or ($target.StartsWith('<') -and -not $target.StartsWith('<<'))) { $token = $target }
                }
            }
            if (-not $token) { continue }
            $text = ConvertFrom-PdfTextString $token
            if ($check.Key -eq 'RC') { $text = $text -replace '<[^>]+>', ' ' }   # rich text: XHTML tags dropped
            $label = $check.Label
            if ($check.Key -eq 'V') { $name = $nameRegex.Match($obj.Dict); $label = ''; if ($name.Success) { $label = ConvertFrom-PdfTextString $name.Groups[1].Value } }
            if ($label) { $label + ': ' + $text } else { $text }
        }
    }
}

function Get-PdfText {
    <#
    .SYNOPSIS
      The text of a PDF: the text of every page and the values of its form fields. A PDF that is encrypted, has
      no text (a scanned image) or is not a PDF gets a reason instead. Fonts without a ToUnicode map and other
      unusual encodings can give partial text: what can be read is checked.
    .PARAMETER Bytes
      The file bytes.
    .OUTPUTS
      [hashtable] Text, Reason (empty when the text was read), Note (a part not read, for a "check by hand" row; or empty).
    #>
    param([byte[]]$Bytes)
    $text = $script:PdfLatin1.GetString($Bytes)
    if ($text.IndexOf('%PDF') -lt 0 -or $text.IndexOf('%PDF') -gt 1024) {
        return @{ Text = ''; Reason = 'Not a valid PDF (damaged, or another type with a .pdf name): check by hand' }
    }
    if ($text -match '/Encrypt\s*(?:\d+\s+\d+\s+R|<<)') { return @{ Text = ''; Reason = $script:PdfEncryptedReason } }
    $script:PdfScanImages = @()   # this PDF's own scanned pages only
    Start-PdfBudget
    try { $all = Get-PdfAllText $text }
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.TimeoutException] -or $e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
            if ($e -is [System.FormatException] -and $e.Message -eq $script:PdfStringError) { return @{ Text = ''; Reason = $script:PdfPartReason } }
        }
        throw   # any other error is the caller's (Get-TextFromBytes reports the file)
    }
    finally { $notes = @(Stop-PdfBudget) }
    $note = (@($notes) -join '; ')
    if (-not $all.Trim()) {
        if ($script:PdfScanImages.Count -gt 0) { return (Get-PdfScanText $script:PdfScanImages $note) }
        $reason = $script:PdfNoTextReason; if ($note) { $reason = $note }
        return @{ Text = ''; Reason = $reason; Note = '' }
    }
    return @{ Text = $all; Reason = ''; Note = $note }
}

function Get-PdfScanText {
    <#
    .SYNOPSIS
      The text of a scanned PDF (JPEG page images, no text layer) read with Windows OCR, page by page, at most
      $MaxScanPages pages; more pages, or OCR that is not available, give a reason or a note.
    .PARAMETER Images
      The JPEG image streams (Latin-1 text).
    .PARAMETER Note
      Notes already made for this PDF.
    .OUTPUTS
      [hashtable] Text, Reason, Note.
    #>
    param([string[]]$Images, [string]$Note)
    $pages = New-Object System.Collections.Generic.List[string]
    $notes = @($Note | Where-Object { $_ })
    for ($i = 0; $i -lt [Math]::Min($Images.Count, $script:MaxScanPages); $i++) {
        $read = Get-ImageText ($script:PdfLatin1.GetBytes($Images[$i]))
        if ($read.Reason) { return @{ Text = ''; Reason = 'Scanned PDF: ' + $read.Reason; Note = '' } }
        $pages.Add($read.Text)
    }
    if ($Images.Count -gt $script:MaxScanPages) { $notes += "Scanned PDF: only the first $script:MaxScanPages page images were read: check by hand" }
    $text = $pages -join "`n"
    if (-not $text.Trim()) { return @{ Text = ''; Reason = $script:PdfNoTextReason; Note = '' } }
    return @{ Text = $text; Reason = ''; Note = ($notes -join '; ') }
}

function Get-PdfPageResources {
    <#
    .SYNOPSIS
      The resources dictionary of a page: its own (inline or by reference), else the nearest one up its /Parent
      chain (a Pages node), so a page that inherits its fonts gets the right ones.
    .PARAMETER Page
      The page object.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .OUTPUTS
      [string] the dictionary text to look up fonts in.
    #>
    param($Page, $Objects)
    $node = $Page
    for ($level = 0; $level -lt 32 -and $node; $level++) {   # 32: deeper than any real page tree; stops a loop
        $ref = [regex]::Match($node.Dict, '/Resources\s+(\d+)\s+\d+\s+R')
        if ($ref.Success -and $Objects.Contains($ref.Groups[1].Value)) { return $Objects[$ref.Groups[1].Value].Dict }
        if ($node.Dict -match '/Resources\s*<<') { return $node.Dict }
        $parent = [regex]::Match($node.Dict, '/Parent\s+(\d+)\s+\d+\s+R')
        $node = $null
        if ($parent.Success -and $Objects.Contains($parent.Groups[1].Value)) { $node = $Objects[$parent.Groups[1].Value] }
    }
    return $Page.Dict
}

function Get-PdfAllText {
    <#
    .SYNOPSIS
      Page text (each page with its own fonts; the fonts of the whole file when a page names none, as when its
      resources are inherited), the text of form XObjects, and text in dictionaries (Get-PdfDictText), joined.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [string]
    #>
    param([string]$Text)
    $objects = Get-PdfObjects $Text
    $script:PdfScanImages = @($objects.Values | Where-Object { $_.Image } | ForEach-Object { $_.Image })
    Assert-PdfTime
    $allFonts = Get-PdfFonts $objects
    Assert-PdfTime
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($page in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/Page(?![s\w])' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $page $objects) $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $contents = [regex]::Match($page.Dict, '/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)').Groups[1].Value
        foreach ($content in [regex]::Matches($contents, '(\d+)\s+\d+\s+R')) {
            $stream = $objects[$content.Groups[1].Value]
            if (-not $stream) { continue }
            if ($stream.Stream) { $parts.Add((Convert-PdfContentToText $stream.Stream $fonts)); continue }
            # /Contents 5 0 R where object 5 is an array of streams ([6 0 R 7 0 R]): read each of them.
            foreach ($inner in [regex]::Matches($stream.Dict, '(\d+)\s+\d+\s+R')) {
                $part = $objects[$inner.Groups[1].Value]
                if ($part -and $part.Stream) { $parts.Add((Convert-PdfContentToText $part.Stream $fonts)) }
            }
        }
    }
    # Form XObjects (/Subtype /Form, drawn with "Do"): reusable page parts that can hold text of their own. Read each
    # once; a form field's appearance can repeat its value, so a hit may show twice (better than a miss).
    foreach ($form in @($objects.Values | Where-Object { $_.Stream -and $_.Dict -match '/Subtype\s*/Form\b' })) {
        $fonts = Get-PdfFontTable $form.Dict $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $parts.Add((Convert-PdfContentToText $form.Stream $fonts))
    }
    foreach ($line in @(Get-PdfDictText $objects)) { $parts.Add($line) }
    return ($parts -join "`n")
}

Export-ModuleMember -Function Get-PdfText
