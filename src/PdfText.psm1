<#
PdfText.psm1 - the text of a PDF file, in memory, with .NET only (no module; the PAM, 2026-11-20).

  Purpose:     Page text (Tj/TJ text in Flate, ASCII85 or hex streams, with ToUnicode font maps), filled-in form-field
               values, comments and file info, so the detectors can check a PDF (ISO 32000). An encrypted PDF,
               a PDF with no text (a scanned image) or a file that is not a PDF gets a reason instead.
  Inputs:      the file bytes.
  Outputs:     a hashtable with Text and Reason (Get-PdfText).
  How to run:  imported by Find-ChannelPII.ps1 (used by TextExtract.psm1); not run on its own.
  Depends on:  PdfCore.psm1 (limits, stream filters, ToUnicode maps) and PdfObjects.psm1 (objects, strings,
               dictionary text), imported first by Find-ChannelPII.ps1.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression


# Bytes as characters one to one (ISO-8859-1), so binary PDF parts can be cut with string functions.
$script:PdfLatin1 = [System.Text.Encoding]::GetEncoding(28591)

# The tokens of a content stream that matter for text: strings, array brackets, names, numbers and operators.
$script:PdfTokenPattern = (Get-PdfStringPattern) + '|%[^\r\n]*|\[|\]|/[^\s/\[\]()<>{}%]+|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z''"*]+'

# A TJ gap wider than this (thousandths of a text unit) is a space between words.
$script:PdfWordGap = 250

# A same-line Td move leaves a word gap when it starts the next piece more than this share of the font size past
# the estimated end of the text before it (Add-PdfMove); a smaller move (kerning, glyph by glyph) joins the pieces.
$script:WordGapEms = 0.25

# The end of an inline image's bytes (BI <keys> ID <bytes> EI): white space, EI, then a delimiter or the end. After a
# BI operator (a token, never the word BI in a string) the reader steps over the bytes to here, so image bytes never
# join with page text into one token (fresh reviews, 2026-11-21).
$script:InlineImageEnd = '\sEI(?=[\s/\[<(%]|$)'

$script:PdfNoTextReason = 'No text in this PDF (a scanned image?): check by hand'
$script:PdfSlowReason = 'This PDF is too large or damaged to read in time: check by hand'
$script:PdfPartReason = 'Part of this PDF could not be read (a damaged or unusual string): check by hand'
$script:PdfEncryptedReason = 'Encrypted PDF (password or label): check by hand'

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
                elseif ($entry.Dict -match '/Subtype\s*/Type0\b') {
                    # A composite font without a ToUnicode map: its codes are glyph numbers, not letters. Text shown in it
                    # cannot be read: a note (ConvertFrom-PdfGlyphs), never garbage read as text.
                    $entry['Map'] = @{ Bytes = 2; Map = @{}; Ranges = @(); Unreadable = $true }
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
    if ($Font.ContainsKey('Unreadable')) {
        if ($Raw.Length -gt 0) { Add-PdfNote 'A font in this PDF could not be read: check by hand' }
        return ''
    }
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
        if ($null -ne $text) { [void]$out.Append($text) }
        elseif ($step -eq 1) { [void]$out.Append($Raw[$i]) }
        elseif ($number -ne 0) { Add-PdfNote 'Some characters in this PDF could not be read (a font without them in its map): check by hand' }
    }
    return $out.ToString()
}

function Convert-PdfContentToText {
    <#
    .SYNOPSIS
      The text a content stream shows: Tj, TJ, ' and " strings in the current font; a line break at each line move
      and text block end; a space at a wide TJ gap and at a same-line Td move that leaves a gap (Add-PdfMove), so
      text placed piece by piece (Chrome splits at kerning pairs and draws glyph by glyph) joins again.
    .PARAMETER Content
      The (inflated) content stream.
    .PARAMETER Fonts
      Output of Get-PdfFonts or Get-PdfFontTable.
    .OUTPUTS
      [string]
    #>
    param([string]$Content, [hashtable]$Fonts)
    $state = @{ Out = (New-Object System.Text.StringBuilder); Block = (New-Object System.Text.StringBuilder); Font = $null; Size = 12.0; Advance = 0.0 }
    $stack = New-Object System.Collections.Generic.List[object]
    $array = $null
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
        Invoke-PdfTextOperator $t $stack $state $Fonts
        $stack.Clear()
    }
    Assert-PdfGap $Content $next $Content.Length
    Complete-PdfTextBlock $state
    return $state.Out.ToString()
}

function Invoke-PdfTextOperator {
    <#
    .SYNOPSIS
      Applies one content-stream operator to the text state: font change, shown text, line and word breaks.
    .PARAMETER Operator
      The operator token.
    .PARAMETER Stack
      Its operands.
    .PARAMETER State
      Out, Block (the text block in work), Font, Size (font size), Advance (estimated width shown since the line's start).
    .PARAMETER Fonts
      The font table.
    #>
    param([string]$Operator, $Stack, [hashtable]$State, [hashtable]$Fonts)
    $last = $null; if ($Stack.Count -gt 0) { $last = $Stack[$Stack.Count - 1] }
    $shown = $null   # the string a Tj, ' or " shows: only a string token counts
    if ($Operator -cin @('Tj', "'", '"') -and $last -is [string] -and $last.Length -ge 2 -and ($last[0] -eq '(' -or $last[0] -eq '<')) {
        $shown = ConvertFrom-PdfGlyphs (ConvertFrom-PdfString $last) $State.Font
        $State.Advance += Get-PdfTextWidth $shown $State.Size
    }
    $block = $State.Block
    switch -CaseSensitive ($Operator) {
        'Tf' { if ($Stack.Count -ge 2) { $State.Font = $Fonts[([string]$Stack[$Stack.Count - 2]).TrimStart('/')]; $State.Size = [Math]::Abs((ConvertTo-PdfNumber $last)) } }
        'Tj' { [void]$block.Append($shown) }
        "'" { [void]$block.Append("`n" + $shown); $State.Advance = Get-PdfTextWidth $shown $State.Size }
        '"' { [void]$block.Append("`n" + $shown); $State.Advance = Get-PdfTextWidth $shown $State.Size }
        'TJ' {
            if ($last -is [System.Collections.Generic.List[string]]) {
                $text = Convert-PdfArrayToText $last $State.Font
                [void]$block.Append($text); $State.Advance += Get-PdfTextWidth $text $State.Size
            }
        }
        'Td' { Add-PdfMove $Stack $State }
        'TD' { [void]$block.Append("`n"); $State.Advance = 0.0 }
        'T*' { [void]$block.Append("`n"); $State.Advance = 0.0 }
        'Tm' { [void]$block.Append("`n"); $State.Advance = 0.0 }
        'ET' { Complete-PdfTextBlock $State }
    }
}

function Get-PdfTextWidth {
    <#
    .SYNOPSIS
      The estimated width of shown text, in text space units: average glyph widths (Helvetica-like, in ems) times
      the font size. Digits and common number punctuation have their usual widths, so the gap after a number piece
      is measured well.
    .PARAMETER Text
      The shown text.
    .PARAMETER Size
      The font size.
    .OUTPUTS
      [double]
    #>
    param([string]$Text, [double]$Size)
    $ems = 0.0
    foreach ($c in $Text.ToCharArray()) {
        if ([char]::IsDigit($c)) { $ems += 0.556 }
        elseif ($c -eq ' ' -or $c -eq '.' -or $c -eq ',' -or $c -eq ':' -or $c -eq ';') { $ems += 0.278 }
        elseif ($c -eq '-' -or $c -eq '(' -or $c -eq ')' -or $c -eq '/') { $ems += 0.333 }
        elseif ([char]::IsUpper($c)) { $ems += 0.667 }
        else { $ems += 0.5 }
    }
    return ($ems * $Size)
}

function Add-PdfMove {
    <#
    .SYNOPSIS
      A Td move: down or up starts a new line; along the line, a space only where the move leaves a gap after the
      text shown since the line's start (more than $WordGapEms of the font size), else the pieces join.
    .PARAMETER Stack
      The operands (tx ty).
    .PARAMETER State
      The text state.
    #>
    param($Stack, [hashtable]$State)
    if ($Stack.Count -lt 2) { return }
    $tx = ConvertTo-PdfNumber $Stack[$Stack.Count - 2]
    $ty = ConvertTo-PdfNumber $Stack[$Stack.Count - 1]
    if ($ty -ne 0) { [void]$State.Block.Append("`n") }
    elseif ($tx -lt 0 -or $tx - $State.Advance -gt $script:WordGapEms * [Math]::Max(1.0, $State.Size)) { [void]$State.Block.Append(' ') }
    $State.Advance = 0.0
}

function Complete-PdfTextBlock {
    <#
    .SYNOPSIS
      Ends a text block: symbol-font characters back to their plain ones, other private-use characters a space and a
      note; then a line break.
    .PARAMETER State
      The text state of Convert-PdfContentToText.
    #>
    param([hashtable]$State)
    $text = $State.Block.ToString()
    # Microsoft symbol fonts map their characters to U+F020-U+F0FF (Word's bullet is U+F0B7): back to U+0020-U+00FF.
    $text = [regex]::Replace($text, '[\uF020-\uF0FF]', [System.Text.RegularExpressions.MatchEvaluator] { param($m) [string][char]([int]$m.Value[0] - 0xF000) })
    # A font can map a glyph to a private-use character (U+E000-U+F8FF): it has no meaning (Chrome does this for some
    # punctuation), so it becomes a space and the PDF gets a note.
    if ($text -match '[\uE000-\uF8FF]') {
        Add-PdfNote 'Some characters in this PDF could not be read (a font without them in its map): check by hand'
        $text = $text -replace '[\uE000-\uF8FF]', ' '
    }
    [void]$State.Out.Append($text).Append("`n")
    [void]$State.Block.Clear()
    $State.Advance = 0.0
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
    Reset-PdfOldObjects   # this PDF's own older versions only
    Start-PdfBudget
    try { $all = Get-PdfAllText $text }
    catch {
        # A limit or an unread string can arrive wrapped (a .NET call, the pipeline): look through the whole chain.
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.TimeoutException] -or $e -is [System.Text.RegularExpressions.RegexMatchTimeoutException]) { return @{ Text = ''; Reason = $script:PdfSlowReason } }
            if ($e -is [System.FormatException] -and $e.Message -eq (Get-PdfStringError)) { return @{ Text = ''; Reason = $script:PdfPartReason } }
        }
        throw   # any other error is the caller's (Get-TextFromBytes reports the file)
    }
    finally { $notes = @(Stop-PdfBudget) }
    $note = (@($notes) -join '; ')
    if (-not $all.Trim()) {
        $reason = $script:PdfNoTextReason; if ($note) { $reason = $note }
        return @{ Text = ''; Reason = $reason; Note = '' }
    }
    return @{ Text = $all; Reason = ''; Note = $note }
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

function Get-PdfLooseText {
    <#
    .SYNOPSIS
      The text of content streams no page or form reads: older page versions kept by an incremental update ("SSN"
      edited out of a page is still in the file) and streams of pages since removed. A stream counts as content when
      it has a text block (BT, Tj or TJ, ET) and is not XML. One that cannot be read gives a note, not a failure.
    .PARAMETER Objects
      Output of Get-PdfObjects.
    .PARAMETER Fonts
      The fonts of the whole file.
    .PARAMETER Read
      The streams already read.
    .OUTPUTS
      [string] text, one item per stream.
    #>
    param($Objects, [hashtable]$Fonts, $Read)
    $loose = @($Objects.Values) + (Get-PdfOldObjects)
    foreach ($obj in $loose) {
        if (-not $obj.Stream -or $Read.Contains($obj) -or $obj.Dict -match '/Type\s*/(?:ObjStm|XRef)\b') { continue }
        $start = $obj.Stream.TrimStart()
        if ($start.StartsWith('<') -or $obj.Stream -notmatch '\bBT\b' -or $obj.Stream -notmatch '\bT[Jj]\b|\)\s*[''"]') { continue }
        try { Convert-PdfContentToText $obj.Stream $Fonts }
        catch {
            if ($_.Exception -is [System.FormatException] -and $_.Exception.Message -eq (Get-PdfStringError)) {
                Add-PdfNote 'Part of this PDF could not be read: check by hand'
            }
            else { throw }
        }
    }
}

function Get-PdfAllText {
    <#
    .SYNOPSIS
      Page text (each page with its own fonts; the fonts of the whole file when a page names none, as when its
      resources are inherited), form XObjects, dictionaries (Get-PdfDictText), XFA data and content streams no page
      reads (Get-PdfLooseText), joined.
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
    $read = New-Object System.Collections.Generic.HashSet[object]   # the content streams read as pages or forms
    foreach ($page in @($objects.Values | Where-Object { $_.Dict -match '/Type\s*/Page(?![s\w])' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $page $objects) $objects
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        $contents = [regex]::Match($page.Dict, '/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)').Groups[1].Value
        foreach ($content in [regex]::Matches($contents, '(\d+)\s+\d+\s+R')) {
            $stream = $objects[$content.Groups[1].Value]
            if (-not $stream) { continue }
            if ($stream.Stream) { [void]$read.Add($stream); $parts.Add((Convert-PdfContentToText $stream.Stream $fonts)); continue }
            # /Contents 5 0 R where object 5 is an array of streams ([6 0 R 7 0 R]): read each of them.
            foreach ($inner in [regex]::Matches($stream.Dict, '(\d+)\s+\d+\s+R')) {
                $part = $objects[$inner.Groups[1].Value]
                if ($part -and $part.Stream) { [void]$read.Add($part); $parts.Add((Convert-PdfContentToText $part.Stream $fonts)) }
            }
        }
    }
    # Form XObjects (/Subtype /Form, drawn with "Do"): reusable page parts that can hold text of their own. Read each
    # once; a form field's appearance can repeat its value, so a hit may show twice (better than a miss).
    foreach ($form in @($objects.Values | Where-Object { $_.Stream -and $_.Dict -match '/Subtype\s*/Form\b' })) {
        $fonts = Get-PdfFontTable (Get-PdfPageResources $form $objects) $objects   # also /Resources 10 0 R
        if ($fonts.Count -eq 0) { $fonts = $allFonts }
        [void]$read.Add($form)
        $parts.Add((Convert-PdfContentToText $form.Stream $fonts))
    }
    foreach ($line in @(Get-PdfDictText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfXfaText $objects)) { $parts.Add($line) }
    foreach ($line in @(Get-PdfLooseText $objects $allFonts $read)) { $parts.Add($line) }
    return ($parts -join "`n")
}

Export-ModuleMember -Function Get-PdfText
