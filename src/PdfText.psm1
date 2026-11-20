<#
PdfText.psm1 - the text of a PDF file, in memory, with .NET only (no module; the PAM, 2026-11-20).

  Purpose:     Page text (Tj/TJ text in Flate, ASCII85 or hex streams, with ToUnicode font maps) and filled-in
               form-field values (name: value), so the detectors can check a PDF (ISO 32000). An encrypted PDF,
               a PDF with no text (a scanned image) or a file that is not a PDF gets a reason instead.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text and Reason (Get-PdfText).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary).
$script:PdfStringPattern = '\((?>[^()\\]+|\\.|\((?<d>)|\)(?<-d>))*(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = $script:PdfStringPattern + '|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'

# A TJ gap wider than this (thousandths of a text unit) is a space between words.
$script:PdfWordGap = 250

# A ToUnicode range longer than this is not expanded (a damaged or hostile file cannot fill the memory).
$script:MaxCMapRange = 65536

# Limits, so one damaged or hostile PDF cannot stall the scan (fresh review, 2026-11-20): the time for one PDF, the
# time for one regex search, the size one stream may inflate to, and the inflated size of one PDF in all.
$script:PdfSecondsLimit = 30
$script:PdfRegexTimeout = [TimeSpan]::FromSeconds(2)
$script:MaxStreamBytes = 16MB
$script:MaxInflatedBytes = 64MB
$script:PdfBudget = @{ Clock = $null; Inflated = 0L }

# Streams that hold no page text and are not inflated: images and embedded font programs.
$script:NotTextStream = '/Subtype\s*/(?:Image|Type1C|CIDFontType0C|OpenType)\b|/Length[123]\s'

$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'
$script:PdfSlowReason = 'This PDF is too large or damaged to read in time: check by hand'
$script:PdfEncryptedReason = 'Encrypted PDF (password or label): check by hand'

function Get-PdfRegex {
    # A regex with a time limit per search ($PdfRegexTimeout); a search over it throws instead of running on.
    param([string]$Pattern)
    return New-Object System.Text.RegularExpressions.Regex($Pattern, [System.Text.RegularExpressions.RegexOptions]::None, $script:PdfRegexTimeout)
}

function Assert-PdfTime {
    # Throws when this PDF has taken longer than $PdfSecondsLimit (Get-PdfText then reports it as too slow to read).
    if ($script:PdfBudget.Clock -and $script:PdfBudget.Clock.Elapsed.TotalSeconds -gt $script:PdfSecondsLimit) {
        throw (New-Object System.TimeoutException 'PDF time limit')
    }
}

function Expand-PdfStream {
    <#
    .SYNOPSIS
      A FlateDecode stream (zlib: a 2-byte header, then deflate) inflated; what was read before an error is kept.
    .PARAMETER Raw
      The raw bytes as Latin-1 text.
    #>
    param([string]$Raw)
    $bytes = $script:Latin1.GetBytes($Raw)
    $skip = 0; if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x78) { $skip = 2 }
    $source = New-Object System.IO.MemoryStream -ArgumentList $bytes, $skip, ($bytes.Length - $skip)
    $inflate = New-Object System.IO.Compression.DeflateStream($source, [System.IO.Compression.CompressionMode]::Decompress)
    $output = New-Object System.IO.MemoryStream
    $buffer = New-Object byte[] 65536
    try {
        # Inflate in steps and stop at the caps: a small stream can unpack to gigabytes.
        while ($output.Length -lt $script:MaxStreamBytes -and $script:PdfBudget.Inflated -lt $script:MaxInflatedBytes) {
            $read = $inflate.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $output.Write($buffer, 0, $read)
            $script:PdfBudget.Inflated += $read
        }
    }
    catch { Write-Verbose 'A PDF stream ended early: its text so far is kept.' }
    finally { $inflate.Dispose(); $source.Dispose() }
    return $script:Latin1.GetString($output.ToArray())
}

function Read-PdfStream {
    <#
    .SYNOPSIS
      A stream's data with its filters undone in order: FlateDecode, ASCII85Decode, ASCIIHexDecode (often chained,
      such as [/ASCII85Decode /FlateDecode]). $null for any other filter (images and the like are not text).
    .PARAMETER Dict
      The stream's dictionary.
    .PARAMETER Raw
      The stream's bytes as Latin-1 text.
    .OUTPUTS
      [string] or $null.
    #>
    param([string]$Dict, [string]$Raw)
    if ($Dict -match $script:NotTextStream) { return $null }
    $filter = [regex]::Match($Dict, '/Filter\s*(\[[^\]]*\]|/\w+)')
    $data = $Raw
    foreach ($name in @([regex]::Matches($filter.Groups[1].Value, '/(\w+)') | ForEach-Object { $_.Groups[1].Value })) {
        switch ($name) {
            'FlateDecode' { $data = Expand-PdfStream $data }
            'ASCII85Decode' { $data = ConvertFrom-Ascii85 $data }
            'ASCIIHexDecode' { $data = ConvertFrom-PdfString ('<' + ($data -replace '[^0-9A-Fa-f]', '') + '>') }
            default { return $null }
        }
    }
    return $data
}

function ConvertFrom-Ascii85 {
    <#
    .SYNOPSIS
      ASCII85 text (PDF ASCII85Decode, ending in ~>) to its bytes as Latin-1 text; z = four zero bytes.
    .PARAMETER Text
      The text.
    #>
    param([string]$Text)
    $end = $Text.IndexOf('~>'); if ($end -ge 0) { $Text = $Text.Substring(0, $end) }
    $Text = ($Text -replace '\s', '') -replace 'z', '!!!!!'
    $out = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Text.Length; $i += 5) {
        $group = $Text.Substring($i, [Math]::Min(5, $Text.Length - $i))
        $count = $group.Length - 1
        if ($count -lt 1) { break }
        $group = $group.PadRight(5, 'u')
        [long]$value = 0
        foreach ($c in $group.ToCharArray()) { $value = $value * 85 + ([int]$c - 33) }
        for ($k = 0; $k -lt $count; $k++) { [void]$out.Append([char](($value -shr (24 - 8 * $k)) -band 0xFF)) }
    }
    return $out.ToString()
}

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
        if ($parts.Success) { $body = $parts.Groups[1].Value; $stream = Read-PdfStream $body $parts.Groups[2].Value }
        $objects[$heads[$h].Groups[1].Value] = @{ Dict = $body; Stream = $stream }
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
    if ($first -le 0 -or $first -gt $Packed.Stream.Length) { return }
    $pairs = @($Packed.Stream.Substring(0, $first).Trim() -split '\s+')
    $number = 0
    if (@($pairs | Where-Object { -not [int]::TryParse($_, [ref]$number) }).Count -gt 0) { return }   # a damaged list: skip it
    for ($k = 0; $k -lt $count -and 2 * $k + 1 -lt $pairs.Count; $k++) {
        $start = $first + [int]$pairs[2 * $k + 1]
        $end = $Packed.Stream.Length
        if (2 * $k + 3 -lt $pairs.Count) { $end = $first + [int]$pairs[2 * $k + 3] }
        if ($start -ge $end -or $end -gt $Packed.Stream.Length) { continue }
        if (-not $Objects.Contains($pairs[2 * $k])) { $Objects[$pairs[2 * $k]] = @{ Dict = $Packed.Stream.Substring($start, $end - $start); Stream = $null } }
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
        return -join @(for ($i = 0; $i -lt $hex.Length; $i += 2) { [char][Convert]::ToInt32($hex.Substring($i, 2), 16) })
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
        return [System.Text.Encoding]::BigEndianUnicode.GetString($script:Latin1.GetBytes($raw.Substring(2)))
    }
    return $raw
}

function ConvertFrom-Utf16Hex {
    # The text of a ToUnicode target: UTF-16BE given as hex.
    param([string]$Hex)
    $bytes = [byte[]]@(for ($i = 0; $i + 1 -lt $Hex.Length; $i += 2) { [Convert]::ToByte($Hex.Substring($i, 2), 16) })
    return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
}

function Get-PdfCMap {
    <#
    .SYNOPSIS
      A font's ToUnicode map: glyph code (hex, upper case) -> text, from its bfchar and bfrange entries.
    .PARAMETER Text
      The CMap stream.
    .OUTPUTS
      [hashtable] Bytes (bytes per code), Map.
    #>
    param([string]$Text)
    $map = @{}
    $bytes = 1
    foreach ($block in [regex]::Matches($Text, '(?s)beginbfchar(.*?)endbfchar')) {
        foreach ($pair in [regex]::Matches($block.Groups[1].Value, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>')) {
            $map[$pair.Groups[1].Value.ToUpper()] = ConvertFrom-Utf16Hex $pair.Groups[2].Value
            $bytes = [int]($pair.Groups[1].Value.Length / 2)
        }
    }
    foreach ($block in [regex]::Matches($Text, '(?s)beginbfrange(.*?)endbfrange')) {
        foreach ($range in [regex]::Matches($block.Groups[1].Value, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])')) {
            $width = $range.Groups[1].Value.Length
            $bytes = [int]($width / 2)
            $low = [Convert]::ToInt64($range.Groups[1].Value, 16)
            $high = [Convert]::ToInt64($range.Groups[2].Value, 16)
            if ($high -lt $low -or $high - $low -ge $script:MaxCMapRange) { continue }
            $target = $range.Groups[3].Value
            $list = @([regex]::Matches($target, '<([0-9A-Fa-f]*)>') | ForEach-Object { $_.Groups[1].Value })
            for ($code = $low; $code -le $high; $code++) {
                $key = $code.ToString('X' + $width)
                if ($target.StartsWith('[')) { if ($code - $low -lt $list.Count) { $map[$key] = ConvertFrom-Utf16Hex $list[$code - $low] } }
                elseif ($list.Count -gt 0 -and $list[0].Length -ge 4) {
                    $last = [Convert]::ToInt32($list[0].Substring($list[0].Length - 4), 16) + ($code - $low)
                    $map[$key] = ConvertFrom-Utf16Hex ($list[0].Substring(0, $list[0].Length - 4) + ($last -band 0xFFFF).ToString('X4'))
                }
            }
        }
    }
    return @{ Bytes = $bytes; Map = $map }
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
                    $entry['Map'] = Get-PdfCMap $Objects[$unicode.Groups[1].Value].Stream
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
        $code = -join @(for ($j = 0; $j -lt $step; $j++) { ([int]$Raw[$i + $j]).ToString('X2') })
        $text = $Font.Map[$code]
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
    $inImage = $false   # between ID and EI: the bytes of an inline image, not text
    $count = 0
    foreach ($token in (Get-PdfRegex $script:PdfTokenPattern).Matches($Content)) {
        if ((++$count % 2000) -eq 0) { Assert-PdfTime }
        $t = $token.Value
        if ($inImage) { if ($t -ceq 'EI') { $inImage = $false }; continue }
        if ($t -ceq 'ID') { $inImage = $true; $stack.Clear(); continue }
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
    return $out.ToString()
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

function Get-PdfFormValues {
    <#
    .SYNOPSIS
      The filled-in form fields: one line "name: value" per field with a text value (/V), so labels stay with values.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    #>
    param($Objects)
    $valueRegex = Get-PdfRegex ('/V\s*(' + $script:PdfStringPattern + ')')
    $nameRegex = Get-PdfRegex ('/T\s*(' + $script:PdfStringPattern + ')')
    foreach ($obj in $Objects.Values) {
        Assert-PdfTime
        if ($obj.Dict.IndexOf('/V') -lt 0) { continue }
        $value = $valueRegex.Match($obj.Dict)
        if (-not $value.Success) { continue }
        $name = $nameRegex.Match($obj.Dict)
        $label = ''; if ($name.Success) { $label = (ConvertFrom-PdfTextString $name.Groups[1].Value) + ': ' }
        $label + (ConvertFrom-PdfTextString $value.Groups[1].Value)
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
      [hashtable] Text, Reason (empty when the text was read).
    #>
    param([byte[]]$Bytes)
    $text = $script:Latin1.GetString($Bytes)
    if ($text.IndexOf('%PDF') -lt 0 -or $text.IndexOf('%PDF') -gt 1024) {
        return @{ Text = ''; Reason = 'Not a valid PDF (damaged, or another type with a .pdf name): check by hand' }
    }
    if ($text -match '/Encrypt\s*(?:\d+\s+\d+\s+R|<<)') { return @{ Text = ''; Reason = $script:PdfEncryptedReason } }
    $script:PdfBudget = @{ Clock = [Diagnostics.Stopwatch]::StartNew(); Inflated = 0L }
    try { $all = Get-PdfAllText $text }
    catch [System.TimeoutException], [System.Text.RegularExpressions.RegexMatchTimeoutException] { return @{ Text = ''; Reason = $script:PdfSlowReason } }
    catch {
        # A timeout raised inside a .NET call arrives wrapped; any other error is the caller's (Get-TextFromBytes).
        if ($_.Exception.InnerException -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
        throw
    }
    finally { $script:PdfBudget = @{ Clock = $null; Inflated = 0L } }
    if (-not $all.Trim()) { return @{ Text = ''; Reason = $script:PdfNoTextReason } }
    return @{ Text = $all; Reason = '' }
}

function Get-PdfAllText {
    <#
    .SYNOPSIS
      Page text (each page with its own fonts; the fonts of the whole file when a page names none, as when its
      resources are inherited) and form-field values, joined with line breaks.
    .PARAMETER Text
      The PDF bytes as Latin-1 text.
    .OUTPUTS
      [string]
    #>
    param([string]$Text)
    $objects = Get-PdfObjects $Text
    $allFonts = Get-PdfFonts $objects
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($page in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/Page(?![s\w])' })) {
        $resources = $page.Dict
        $ref = [regex]::Match($page.Dict, '/Resources\s+(\d+)\s+\d+\s+R')
        if ($ref.Success -and $objects.Contains($ref.Groups[1].Value)) { $resources = $objects[$ref.Groups[1].Value].Dict }
        $fonts = Get-PdfFontTable $resources $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $contents = [regex]::Match($page.Dict, '/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)').Groups[1].Value
        foreach ($content in [regex]::Matches($contents, '(\d+)\s+\d+\s+R')) {
            $stream = $objects[$content.Groups[1].Value]
            if ($stream -and $stream.Stream) { $parts.Add((Convert-PdfContentToText $stream.Stream $fonts)) }
        }
    }
    foreach ($line in @(Get-PdfFormValues $objects)) { $parts.Add($line) }
    return ($parts -join "`n")
}

Export-ModuleMember -Function Get-PdfText
