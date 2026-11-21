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

# A PDF string: (literal), with nested brackets and escapes, or <hex> (not a << dictionary). A literal holds at most
# 10,000 parts, so an unclosed bracket fails after a bounded scan. A "(" that no string token takes means a string
# was not read: the PDF is then NOT SCANNED, never read in part without a word (Assert-PdfGap).
$script:PdfStringPattern = '\((?>[^()\\]+|\\.|\((?<d>)|\)(?<-d>)){0,10000}(?(d)(?!))\)|<(?!<)[0-9A-Fa-f\s]*>'

# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = $script:PdfStringPattern + '|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'

# A TJ gap wider than this (thousandths of a text unit) is a space between words.
$script:PdfWordGap = 250

# A ToUnicode map uses at most this many ranges; ranges are looked up, not expanded, so a hostile map costs no time.
$script:MaxCMapRanges = 2000

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
$script:PdfPartReason = 'Part of this PDF could not be read (a damaged or unusual string): check by hand'
$script:PdfStringError = 'PDF string not read'
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
        if (($i % 100000) -eq 0) { Assert-PdfTime }
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

function Get-PdfBlocks {
    <#
    .SYNOPSIS
      The text between each begin word and the end word after it (bfchar or bfrange blocks), found in one linear pass:
      a block without its end word runs to the next begin word, so a damaged map cannot make the search quadratic.
    .PARAMETER Text
      The CMap stream.
    .PARAMETER Begin
      The begin word, for example 'beginbfchar'.
    .PARAMETER End
      The end word, for example 'endbfchar'.
    .OUTPUTS
      [string] one per block.
    #>
    param([string]$Text, [string]$Begin, [string]$End)
    $starts = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($Begin))).Matches($Text)) { $m.Index + $m.Length })
    $stops = @(foreach ($m in (Get-PdfRegex ([regex]::Escape($End))).Matches($Text)) { $m.Index })
    $e = 0
    for ($b = 0; $b -lt $starts.Count; $b++) {
        $stop = $Text.Length; if ($b + 1 -lt $starts.Count) { $stop = $starts[$b + 1] - $Begin.Length }
        while ($e -lt $stops.Count -and $stops[$e] -lt $starts[$b]) { $e++ }
        if ($e -lt $stops.Count -and $stops[$e] -lt $stop) { $stop = $stops[$e] }
        $Text.Substring($starts[$b], $stop - $starts[$b])
    }
}

function Get-PdfCMap {
    <#
    .SYNOPSIS
      A font's ToUnicode map: glyph code (hex, upper case) -> text from its bfchar entries and listed ranges, plus
      its plain ranges (looked up by ConvertFrom-PdfGlyphs, not expanded).
    .PARAMETER Text
      The CMap stream.
    .OUTPUTS
      [hashtable] Bytes (bytes per code), Map, Ranges.
    #>
    param([string]$Text)
    $map = @{}
    $ranges = New-Object System.Collections.Generic.List[object]
    $bytes = 1
    $pairRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>'
    $rangeRegex = Get-PdfRegex '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]*>|\[[^\]]*\])'
    $listRegex = Get-PdfRegex '<([0-9A-Fa-f]*)>'
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfchar' 'endbfchar')) {
        Assert-PdfTime
        foreach ($pair in $pairRegex.Matches($block)) {
            $map[$pair.Groups[1].Value.ToUpper()] = ConvertFrom-Utf16Hex $pair.Groups[2].Value
            $bytes = [int]($pair.Groups[1].Value.Length / 2)
        }
    }
    foreach ($block in @(Get-PdfBlocks $Text 'beginbfrange' 'endbfrange')) {
        Assert-PdfTime
        foreach ($range in $rangeRegex.Matches($block)) {
            $width = $range.Groups[1].Value.Length
            $bytes = [int]($width / 2)
            $low = [Convert]::ToInt64($range.Groups[1].Value, 16)
            $high = [Convert]::ToInt64($range.Groups[2].Value, 16)
            $target = $range.Groups[3].Value
            if ($high -lt $low) { continue }
            if ($target.StartsWith('[')) {
                # A listed range names each target: as many entries as the map's own text holds, so no expansion cost.
                $code = $low
                foreach ($item in $listRegex.Matches($target)) {
                    if ($code -gt $high) { break }
                    $map[$code.ToString('X' + $width)] = ConvertFrom-Utf16Hex $item.Groups[1].Value
                    $code++
                }
            }
            elseif ($ranges.Count -lt $script:MaxCMapRanges -and $target.Length -ge 6) {
                $hex = $target.Substring(1, $target.Length - 2)
                $ranges.Add(@{ Low = $low; High = $high; Prefix = $hex.Substring(0, $hex.Length - 4); Start = [Convert]::ToInt32($hex.Substring($hex.Length - 4), 16) })
            }
        }
    }
    return @{ Bytes = $bytes; Map = $map; Ranges = $ranges }
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
    $inImage = $false   # between ID and EI: the bytes of an inline image, not text
    $count = 0
    $next = 0           # where the last token ended: text between tokens must hold no "(" (Assert-PdfGap)
    foreach ($token in (Get-PdfRegex $script:PdfTokenPattern).Matches($Content)) {
        if ((++$count % 100) -eq 0) { Assert-PdfTime }
        $t = $token.Value
        if (-not $inImage) { Assert-PdfGap $Content $next $token.Index }
        $next = $token.Index + $token.Length
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
    if (-not $inImage) { Assert-PdfGap $Content $next $Content.Length }
    return $out.ToString()
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
        if (-not $value.Success) {
            if ($obj.Dict -match '/V\s*\(') { throw (New-Object System.FormatException $script:PdfStringError) }   # a value not read
            continue
        }
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
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.TimeoutException] -or $e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
            if ($e -is [System.FormatException] -and $e.Message -eq $script:PdfStringError) { return @{ Text = ''; Reason = $script:PdfPartReason } }
        }
        throw   # any other error is the caller's (Get-TextFromBytes reports the file)
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
    Assert-PdfTime
    $allFonts = Get-PdfFonts $objects
    Assert-PdfTime
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
