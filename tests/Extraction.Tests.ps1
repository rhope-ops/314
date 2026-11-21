<#
Extraction.Tests.ps1 - tests of TextExtract.psm1: text from Office files, text files and HTML.

  Purpose:     Each supported type gives its text; split Word runs join again; other types and damaged
               files are reported as NOT SCANNED.
  Inputs:      $Sample and New-ZipBytes (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal; $Docx and $Txt for the later test files.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- Word: an SSN split over two runs must join again ----
$ssnHead = $Sample.Ssn.Substring(0, 7)
$ssnTail = $Sample.Ssn.Substring(7)
$wordXml = "<w:document><w:body><w:p><w:r><w:t>SSN $ssnHead</w:t></w:r><w:r><w:t>$ssnTail</w:t></w:r></w:p>" +
           '<w:p><w:r><w:t>Next &amp; line</w:t></w:r></w:p></w:body></w:document>'
$Docx = New-ZipBytes @{ 'word/document.xml' = $wordXml }
$result = Get-TextFromBytes -Bytes $Docx -FileName 'Roster.docx'
Assert-Equal 'OK' $result.Status 'docx read'
Assert-Equal 'SSN/High/***-**-6789' (Get-HitSummary $result.Text) 'docx split runs joined'
Assert-Equal $true ($result.Text -match 'Next & line') 'docx XML entities decoded'

# ---- Excel: shared strings and cell values ----
$sheetXml = "<worksheet><sheetData><row><c t=`"s`"><v>0</v></c><c><v>$($Sample.DodId)</v></c></row></sheetData></worksheet>"
$xlsx = New-ZipBytes @{ 'xl/sharedStrings.xml' = '<sst><si><t>DoD ID</t></si></sst>'; 'xl/worksheets/sheet1.xml' = $sheetXml }
$result = Get-TextFromBytes -Bytes $xlsx -FileName 'Recall.xlsx'
Assert-Equal 'OK' $result.Status 'xlsx read'
Assert-Equal $true $result.Text.Contains($Sample.DodId) 'xlsx cell value read'
Assert-Equal "DoD ID | $($Sample.DodId)" $result.Text 'xlsx shared string back in its cell, next to its value'

# ---- Excel: no false positives from neighbouring cells, decimals, formulas or string indexes ----
$noise = '<worksheet><sheetData>' +
    '<row><c><v>123</v></c><c><v>45</v></c><c><v>6789</v></c></row>' +                            # three cells, not one SSN
    '<row><c><v>45231.416666666664</v></c><c><v>1.2345678901E+20</v></c></row>' +                  # a time; an exponent
    '<row><c><f>VLOOKUP(A1,Sheet2!A:B,2)*123456789</f><v>7</v></c></row>' +                        # formula text is dropped
    '<row>' + ((1..16 | ForEach-Object { '<c><v>4</v></c>' }) -join '') + '</row>' +                # 16 single digits, not a card
    '<row><c t="b"><v>1</v></c><c t="str"><v>OK</v></c><c t="inlineStr"><is><t>note</t></is></c></row>' +
    '</sheetData></worksheet>'
$noiseText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $noise }) -FileName 'Budget.xlsx').Text
Assert-Equal 0 @(Find-PiiInText $noiseText).Count 'Excel: no hit from cells, decimals, formulas or digit runs'
Assert-Equal $false $noiseText.Contains('VLOOKUP') 'Excel: formula text not read'
Assert-Equal $true ($noiseText -match '7' -and $noiseText -match 'OK \| note') 'Excel: formula results and text cells kept'
# Tags with a namespace prefix (some tools write <x:row>, <x:c>) and rich text in a shared string are read too.
$prefixed = "<x:worksheet><x:sheetData><x:row><x:c t=`"s`"><x:v>0</x:v></x:c><x:c><x:v>$($Sample.SsnBare)</x:v></x:c></x:row></x:sheetData></x:worksheet>"
$richShared = '<x:sst><x:si><x:r><x:t>S</x:t></x:r><x:r><x:rPr><x:b/></x:rPr><x:t>SN</x:t></x:r></x:si></x:sst>'
$prefixedText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $prefixed; 'xl/sharedStrings.xml' = $richShared }) -FileName 'p.xlsx').Text
Assert-Equal "SSN | $($Sample.SsnBare)" $prefixedText 'Excel: prefixed tags and rich-text shared strings read'
Assert-Equal 'SSN' (@(Find-PiiInText $prefixedText) | Select-Object -First 1).Detector 'Excel: SSN found in a file with prefixed tags'
$selfClosing = '<worksheet><sheetData><row><c r="A1"/><c r="B1"><v>5</v></c></row></sheetData></worksheet>'
Assert-Equal '5' (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $selfClosing }) -FileName 's.xlsx').Text 'Excel: an empty self-closing cell is skipped'
$realSsn = "<worksheet><sheetData><row><c t=`"inlineStr`"><is><t>SSN</t></is></c><c><v>$($Sample.SsnBare)</v></c></row></sheetData></worksheet>"
Assert-Equal 1 @(Find-PiiInText (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'xl/worksheets/sheet1.xml' = $realSsn }) -FileName 'a.xlsx').Text |
    Where-Object { $_.Detector -eq 'SSN' }).Count 'Excel: a real SSN in a number cell is still found'

# ---- Word field codes are not read (a TOC bookmark number is no SSN) ----
$fieldXml = "<w:document><w:p><w:r><w:instrText xml:space=`"preserve`"> PAGEREF _Toc$($Sample.SsnBare) \h </w:instrText></w:r><w:r><w:t>Chapter 1</w:t></w:r></w:p></w:document>"
$fieldText = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'word/document.xml' = $fieldXml }) -FileName 'Plan.docx').Text
Assert-Equal 'Chapter 1' $fieldText.Trim() 'docx field code text dropped, page text kept'
$selfClosingField = "<w:document><w:p><w:r><w:instrText/></w:r><w:r><w:t>SSN $($Sample.Ssn)</w:t></w:r><w:r><w:instrText> PAGE </w:instrText></w:r></w:p></w:document>"
$kept = (Get-TextFromBytes -Bytes (New-ZipBytes @{ 'word/document.xml' = $selfClosingField }) -FileName 'a.docx').Text
Assert-Equal 'SSN' (@(Find-PiiInText $kept) | Select-Object -First 1).Detector 'An empty field code does not hide the page text after it'

# ---- PowerPoint ----
$pptx = New-ZipBytes @{ 'ppt/slides/slide1.xml' = '<p:sld><a:p><a:r><a:t>DOB: 03/14/1985</a:t></a:r></a:p></p:sld>' }
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary (Get-TextFromBytes -Bytes $pptx -FileName 'Brief.pptx').Text) 'pptx read'

# ---- Text, HTML, and types that are not read ----
$Txt = [System.Text.Encoding]::UTF8.GetBytes("my cell $($Sample.Phone)")
Assert-Equal 'Phone/Medium/***-***-0142' (Get-HitSummary (Get-TextFromBytes -Bytes $Txt -FileName 'notes.txt').Text) 'txt read'
# ---- PDF: page text (compressed or not, with nested brackets), ToUnicode fonts, form fields, reasons ----
$pdf = New-PdfBytes "BT /F1 12 Tf 72 700 Td (SSN: $($Sample.Ssn)) Tj 0 -14 Td [(my cell \(home\) ) -300 ($($Sample.Phone))] TJ ET"
$pdfText = Get-TextFromBytes -Bytes $pdf -FileName 'Recall.pdf'
Assert-Equal 'OK' $pdfText.Status 'PDF: read'
Assert-Equal 'SSN' (@(Find-PiiInText $pdfText.Text) | Select-Object -First 1).Detector 'PDF: SSN found in compressed page text'
Assert-Equal $true ($pdfText.Text -match 'my cell \(home\)') 'PDF: escaped brackets undone'
$plainPdf = New-PdfBytes "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $plainPdf -FileName 'a.pdf').Text) | Select-Object -First 1).Detector 'PDF: uncompressed page text read'
# A Type0 font with 2-byte glyph codes: the text is only readable through its ToUnicode map.
$cmap = "begincmap 1 begincodespacerange <0000> <FFFF> endcodespacerange 1 beginbfrange <0020> <007E> <0020> endbfrange endcmap"
$glyphs = -join @("SSN: $($Sample.Ssn)".ToCharArray() | ForEach-Object { ([int]$_).ToString('X4') })
$unicodePdf = New-PdfBytes "BT /F1 12 Tf <$glyphs> Tj ET" -ToUnicode $cmap
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $unicodePdf -FileName 'u.pdf').Text) | Select-Object -First 1).Detector 'PDF: ToUnicode font decoded'
$formPdf = New-PdfBytes 'q Q' -Extra @("<< /FT /Tx /T (SSN) /V ($($Sample.Ssn)) >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $formPdf -FileName 'DD Form.pdf').Text) | Select-Object -First 1).Detector 'PDF: form field value found with its field name'
Assert-Equal 'No text in this PDF (a scanned image?): check by hand' (Get-TextFromBytes -Bytes (New-PdfBytes 'q 100 0 0 100 0 0 cm /Im1 Do Q') -FileName 's.pdf').Reason 'PDF: an image-only page named'
Assert-Equal $true ((Get-TextFromBytes -Bytes (New-PdfBytes 'BT (x) Tj ET' -Trailer '/Encrypt 9 0 R') -FileName 'e.pdf').Reason -match '^Encrypted PDF') 'PDF: encrypted named'
Assert-Equal $true ((Get-TextFromBytes -Bytes $Txt -FileName 'scan.pdf').Reason -match '^Not a valid PDF') 'PDF: not a PDF named'
$a85 = & (Get-Module PdfCore) { ConvertFrom-Ascii85 '9jqo^BlbD-BleB1DJ+*+F(f,q~>' }
Assert-Equal 'Man is distinguished' $a85 'PDF: ASCII85 filter decoded (a published test vector)'
$hexData = & (Get-Module PdfCore) { Read-PdfStream '/Filter /ASCIIHexDecode' '48 49>' }
Assert-Equal 'HI' $hexData 'PDF: ASCIIHex filter decoded'
# Object stream (PDF 1.5): a form field packed inside one is still read.
$packed = "9 0 << /T (SSN) /V ($($Sample.Ssn)) >>"
$objStm = New-PdfBytes 'q Q' -Extra @("<< /Type /ObjStm /N 1 /First 4 /Length $($packed.Length) >> stream`n$packed`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $objStm -FileName 'p.pdf').Text) | Select-Object -First 1).Detector 'PDF: object stream unpacked'
# A stream whose first line ends in CR only.
$latin = [System.Text.Encoding]::GetEncoding(28591)
$crPdf = $latin.GetBytes(($latin.GetString((New-PdfBytes "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain)) -replace "stream`n", "stream`r"))
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $crPdf -FileName 'cr.pdf').Text) | Select-Object -First 1).Detector 'PDF: stream line ending in CR read'
# An inline image and a Tj on a number: no error, the text after it still read.
$inline = Get-TextFromBytes -Bytes (New-PdfBytes "BI /W 1 /H 1 ID 7 Tj EI BT 7 Tj (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'i.pdf'
Assert-Equal 'OK|SSN' ($inline.Status + '|' + (@(Find-PiiInText $inline.Text) | Select-Object -First 1).Detector) 'PDF: inline image and Tj on a number skipped'
# Hostile input stays inside the time limit (lowered to 2 s here): unbalanced brackets, and headers without endobj.
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 2 }
$clock = [Diagnostics.Stopwatch]::StartNew()
$slow = Get-TextFromBytes -Bytes (New-PdfBytes ('BT ' + ('(a' * 20000) + ' Tj ET') -Plain) -FileName 'h.pdf'
$noEnd = Get-TextFromBytes -Bytes ([System.Text.Encoding]::ASCII.GetBytes('%PDF-1.4 ' + ('1 0 obj << >> ' * 50000))) -FileName 'n.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15) 'PDF: hostile input ends inside the time limit'
Assert-Equal 'NOT SCANNED|NOT SCANNED' ($slow.Status + '|' + $noEnd.Status) 'PDF: hostile input reported, not read'
# Hostile font maps stay inside the limit: an unclosed bfchar block repeated, and many full ranges (looked up, not expanded).
$clock.Restart()
$badMap = Get-TextFromBytes -Bytes (New-PdfBytes 'BT /F1 12 Tf <0041> Tj ET' -ToUnicode ('beginbfchar ' * 100000)) -FileName 'm.pdf'
$ranges = ('1 beginbfrange <0000> <FFFE> <0000> endbfrange ' * 20)
$manyRanges = Get-TextFromBytes -Bytes (New-PdfBytes ("BT /F1 12 Tf <$glyphs> Tj ET") -ToUnicode $ranges) -FileName 'r.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 15) 'PDF: hostile font maps end inside the time limit'
Assert-Equal $true ($badMap.Status -in @('OK', 'NOT SCANNED')) 'PDF: a damaged map gives a result, not a stall'
Assert-Equal 'SSN' (@(Find-PiiInText $manyRanges.Text) | Select-Object -First 1).Detector 'PDF: full ranges still decode'
& (Get-Module PdfCore) { $script:PdfSecondsLimit = 30 }
# A long string is read; one over the string limit makes the PDF NOT SCANNED, never read in part without a word.
$longString = Get-TextFromBytes -Bytes (New-PdfBytes ('BT (' + ('\101' * 1100) + " $($Sample.Ssn)) Tj ET") -Plain) -FileName 'l.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $longString.Text) | Select-Object -First 1).Detector 'PDF: a string with 1,100 escapes read'
$tooLong = Get-TextFromBytes -Bytes (New-PdfBytes ('BT (hello) Tj (' + ('\101' * 6000) + " $($Sample.Ssn)) Tj ET") -Plain) -FileName 't.pdf'
Assert-Equal $true ($tooLong.Status -eq 'NOT SCANNED' -and $tooLong.Reason -match '^Part of this PDF could not be read') 'PDF: a string over the limit reported, not dropped'
$remarks = Get-TextFromBytes -Bytes (New-PdfBytes 'q Q' -Extra @("<< /T (Remarks) /V (\376\377" + ('\000a' * 1100) + ") >>") -Plain) -FileName 'f.pdf'
Assert-Equal $true ($remarks.Status -eq 'OK' -and $remarks.Text.Contains('a' * 1100)) 'PDF: a long UTF-16 form value read'
# Inline-image bytes cannot hide page text; a bracket in a comment is no unread string.
$imageGap = Get-TextFromBytes -Bytes (New-PdfBytes "BI /W 1 ID xx(yy EI BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET BI /W 1 ID zz)ww EI BT (After) Tj ET" -Plain) -FileName 'g.pdf'
Assert-Equal 'OK|SSN' ($imageGap.Status + '|' + (@(Find-PiiInText $imageGap.Text) | Select-Object -First 1).Detector) 'PDF: inline-image bytes do not hide page text'
$commented = Get-TextFromBytes -Bytes (New-PdfBytes "% draft (rev`nBT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET" -Plain) -FileName 'c.pdf'
Assert-Equal 'OK|SSN' ($commented.Status + '|' + (@(Find-PiiInText $commented.Text) | Select-Object -First 1).Detector) 'PDF: a bracket in a comment is fine'
# Text in a form XObject (a reusable page part drawn with Do) is read.
$xobject = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$formPdf2 = New-PdfBytes '/X1 Do' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Length $($xobject.Length) >> stream`n$xobject`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $formPdf2 -FileName 'x.pdf').Text) | Select-Object -First 1).Detector 'PDF: form XObject text read'
# A part not read is never silent: an attached file, or a stream over the unpack cap, gives a note.
$attached = Get-TextFromBytes -Bytes (New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /EmbeddedFile /Length 3 >> stream`nabc`nendstream")) -FileName 'a.pdf'
Assert-Equal $true ($attached.Status -eq 'OK' -and $attached.Note -match 'attached inside this PDF') 'PDF: attachments give a check-by-hand note'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 10 }
$capped = Get-TextFromBytes -Bytes (New-PdfBytes ("BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET " * 3000)) -FileName 'cap.pdf'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 64MB }
# The cut can end inside a string (then the whole PDF is Not scanned) or between tokens (a note): never silent.
Assert-Equal $true (($capped.Note + $capped.Reason) -match 'too large to unpack|could not be read') 'PDF: text past the unpack cap is never silent'
# A stream that inflates far past the cap is cut, not held whole in memory.
$bomb = & (Get-Module PdfCore) {
    $memory = New-Object System.IO.MemoryStream
    $deflate = New-Object System.IO.Compression.DeflateStream($memory, [System.IO.Compression.CompressionMode]::Compress, $true)
    $zeros = New-Object byte[] (1MB); for ($i = 0; $i -lt 40; $i++) { $deflate.Write($zeros, 0, $zeros.Length) }
    $deflate.Dispose()
    $script:PdfBudget = @{ Clock = $null; Inflated = 0L }
    (Expand-PdfStream ($script:Latin1.GetString($memory.ToArray()))).Length
}
Assert-Equal $true ($bomb -le 16MB + 65536) 'PDF: one stream inflates to at most 16 MB'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $Txt -FileName 'setup.ps1').Status 'script not read'
Assert-Equal 'NOT SCANNED' (Get-TextFromBytes -Bytes $Txt -FileName 'broken.docx').Status 'damaged docx not scanned'
# Bytes that are not a zip get a reason a person can act on, not .NET's "End of central directory not found".
$ole = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0)
Assert-Equal $true ((Get-TextFromBytes -Bytes $ole -FileName 'Recall.xlsx').Reason -match '^Encrypted, password-protected') 'OLE bytes: encrypted or labelled'
Assert-Equal 'Empty file (0 bytes)' (Get-TextFromBytes -Bytes ([byte[]]@()) -FileName 'New.xlsx').Reason 'Empty file named'
$page = [System.Text.Encoding]::ASCII.GetBytes('  <!DOCTYPE html><html>sign in</html>')
Assert-Equal $true ((Get-TextFromBytes -Bytes $page -FileName 'a.xlsx').Reason -match 'web page') 'A web page instead of the file named'
$bomPage = [byte[]](@(0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::ASCII.GetBytes('<html>sign in</html>'))
Assert-Equal $true ((Get-TextFromBytes -Bytes $bomPage -FileName 'a.xlsx').Reason -match 'web page') 'A web page with a byte order mark named'
Assert-Equal $true ((Get-TextFromBytes -Bytes $Txt -FileName 'broken.docx').Reason -match '^Not a valid Office file') 'Other bytes: not a valid Office file'
$html = (Convert-HtmlToText "<div><p>SSN $($Sample.Ssn)</p></div>here") -replace '\s+', ' '
Assert-Equal "SSN $($Sample.Ssn) here" $html.Trim() 'HTML tags removed'

# ---- PDF review round 5 (2026-11-21): no silent loss, bounded time ----
# The word BI in page text is text, not an inline image (an image starts only at a BI operator).
$biText = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (Clearance BI) Tj (ID $($Sample.Ssn)) Tj ( EI) Tj ET" -Plain) -FileName 'bi.pdf'
Assert-Equal 'OK|SSN' ($biText.Status + '|' + (@(Find-PiiInText ("SSN " + $biText.Text)) | Select-Object -First 1).Detector) 'PDF: the word BI does not hide text'
$biPage = Get-TextFromBytes -Bytes (New-PdfBytes ('BT /F1 12 Tf ' + ('(Status: BI complete for member) Tj 0 -14 Td ' * 3000) + 'ET') -Plain) -FileName 'bi2.pdf'
Assert-Equal 'OK' $biPage.Status 'PDF: a long page with the word BI is read in time'
# Comments and sticky notes, a value in an object of its own, and a content list one level deeper.
$annot = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /Annot /Subtype /Text /Contents (SSN $($Sample.Ssn)) >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $annot -FileName 'n.pdf').Text) | Select-Object -First 1).Detector 'PDF: a sticky-note comment read'
$valueRef = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V 8 0 R >>', "($($Sample.Ssn))")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $valueRef -FileName 'v.pdf').Text) | Select-Object -First 1).Detector 'PDF: a form value in an object of its own read'
$inner = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$nested = New-PdfBytes 'q Q' -Plain -PageContents '7 0 R' -Extra @('[8 0 R]', "<< /Length $($inner.Length) >> stream`n$inner`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $nested -FileName 'c.pdf').Text) | Select-Object -First 1).Detector 'PDF: /Contents through an array object read'
# LZW and RunLength streams decoded; any other filter gives a note.
$latin = [System.Text.Encoding]::GetEncoding(28591)
$lzwBits = New-Object System.Text.StringBuilder
foreach ($code in @(256) + @([System.Text.Encoding]::ASCII.GetBytes($inner) | ForEach-Object { [int]$_ }) + @(257)) { [void]$lzwBits.Append([Convert]::ToString($code, 2).PadLeft(9, '0')) }
while ($lzwBits.Length % 8) { [void]$lzwBits.Append('0') }
$lzw = -join @(for ($i = 0; $i -lt $lzwBits.Length; $i += 8) { [char][Convert]::ToInt32($lzwBits.ToString().Substring($i, 8), 2) })
$lzwPdf = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /LZWDecode /Length $($lzw.Length) >> stream`n$lzw`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $lzwPdf -FileName 'l.pdf').Text) | Select-Object -First 1).Detector 'PDF: LZW stream decoded'
$rle = [string][char]($inner.Length - 1) + $inner + [char]128
$rlePdf = New-PdfBytes 'q Q' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /RunLengthDecode /Length $($rle.Length) >> stream`n$rle`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes ($latin.GetBytes($latin.GetString($rlePdf))) -FileName 'r.pdf').Text) | Select-Object -First 1).Detector 'PDF: RunLength stream decoded'
$oddFilter = Get-TextFromBytes -Bytes (New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /XObject /Subtype /Form /Filter /JPXDecode /Length 3 >> stream`nabc`nendstream")) -FileName 'o.pdf'
Assert-Equal $true ($oddFilter.Note -match 'compression this reader cannot open') 'PDF: an unknown filter gives a note'
# A hostile object stream (a huge object list) ends at once with a note.
$clock.Restart()
$hugeList = New-PdfBytes 'BT (hello) Tj ET' -Plain -Extra @("<< /Type /ObjStm /N 200000 /First 4 /Length 8 >> stream`n9 0 (x)`nendstream")
$hugeRead = Get-TextFromBytes -Bytes $hugeList -FileName 'h.pdf'
Assert-Equal $true ($clock.Elapsed.TotalSeconds -lt 10 -and $hugeRead.Note -match 'too large to unpack') 'PDF: a huge object list ends at once with a note'

# ---- PDF review round 6 (2026-11-21) ----
# LZW and RunLength output counts toward the unpack total: past it, a note (bounded memory).
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 10 }
$lzwCapped = Get-TextFromBytes -Bytes $lzwPdf -FileName 'l2.pdf'
& (Get-Module PdfCore) { $script:MaxInflatedBytes = 64MB }
Assert-Equal $true (($lzwCapped.Note + $lzwCapped.Reason) -match 'too large to unpack|could not be read') 'PDF: LZW output counts toward the unpack total'
# A composite font without a ToUnicode map: a reason or note, never glyph numbers read as text.
$type0 = New-PdfBytes 'BT /F1 12 Tf <001400150016> Tj ET' -Plain -FontDict '<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding /Identity-H >>'
$type0Read = Get-TextFromBytes -Bytes $type0 -FileName 't0.pdf'
Assert-Equal $true (($type0Read.Note + $type0Read.Reason) -match 'font in this PDF could not be read') 'PDF: a font without a map is never silent'
# An incremental update: the field defined again later with its value; the old and the new versions are both read.
$filled = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V () >>') -After "7 0 obj << /FT /Tx /T (SSN) /V ($($Sample.Ssn)) >> endobj`n"
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $filled -FileName 'u.pdf').Text) | Select-Object -First 1).Detector 'PDF: a value saved as an added edit is read'
# A multi-select value and a value in a text stream; a backslash before a line end inside a string.
$multi = New-PdfBytes 'q Q' -Plain -Extra @("<< /FT /Ch /T (SSN) /V [(none) ($($Sample.Ssn))] >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $multi -FileName 'm.pdf').Text) | Select-Object -First 1).Detector 'PDF: a multi-select value read'
$vStream = New-PdfBytes 'q Q' -Plain -Extra @('<< /FT /Tx /T (SSN) /V 8 0 R >>', "<< /Length 11 >> stream`n$($Sample.Ssn)`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $vStream -FileName 's2.pdf').Text) | Select-Object -First 1).Detector 'PDF: a value in a text stream read'
$wrapped = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf (SSN: \`n$($Sample.Ssn)) Tj ET" -Plain) -FileName 'w.pdf'
Assert-Equal 'OK|SSN' ($wrapped.Status + '|' + (@(Find-PiiInText $wrapped.Text) | Select-Object -First 1).Detector) 'PDF: a backslash before a line end is a line continuation'

# ---- PDF review round 7 (2026-11-21) ----
# Text drawn one glyph at a time (Chrome): the glyphs join; a normal block keeps its word spaces.
$glyphs1 = (("SSN: $($Sample.Ssn)").ToCharArray() | ForEach-Object { "($_) Tj 7 0 Td" }) -join ' '
$glyphRead = Get-TextFromBytes -Bytes (New-PdfBytes "BT /F1 12 Tf $glyphs1 ET" -Plain) -FileName 'g1.pdf'
Assert-Equal 'SSN' (@(Find-PiiInText $glyphRead.Text) | Select-Object -First 1).Detector 'PDF: text drawn glyph by glyph joins'
$words = Get-TextFromBytes -Bytes (New-PdfBytes 'BT /F1 12 Tf (Member) Tj 60 0 Td (roster) Tj ET' -Plain) -FileName 'w2.pdf'
Assert-Equal 'Member roster' $words.Text.Trim() 'PDF: words placed apart keep their space'
# XFA form data; an older page version kept by an incremental update; two comments in one dictionary.
$xfa = "<xfa:datasets><xfa:data><form><SSN>$($Sample.Ssn)</SSN></form></xfa:data></xfa:datasets>"
$xfaPdf = New-PdfBytes 'BT (Please wait) Tj ET' -Plain -Extra @("<< /Fields [] /XFA 8 0 R >>", "<< /Length $($xfa.Length) >> stream`n$xfa`nendstream")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $xfaPdf -FileName 'x2.pdf').Text) | Select-Object -First 1).Detector 'PDF: XFA form data read'
$oldPage = "BT /F1 12 Tf (SSN: $($Sample.Ssn)) Tj ET"
$edited = New-PdfBytes $oldPage -Plain -After "4 0 obj << /Length 22 >> stream`nBT (SSN: removed) Tj ET`nendstream endobj`n"
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $edited -FileName 'e2.pdf').Text) | Select-Object -First 1).Detector 'PDF: text of an older page version read'
$twoNotes = New-PdfBytes 'q Q' -Plain -Extra @("<< /Annots [<< /Subtype /Text /Contents (first note) >> << /Subtype /Text /Contents (SSN $($Sample.Ssn)) >>] >>")
Assert-Equal 'SSN' (@(Find-PiiInText (Get-TextFromBytes -Bytes $twoNotes -FileName 'n2.pdf').Text) | Select-Object -First 1).Detector 'PDF: every comment in one dictionary read'
# A glyph mapped to a private-use character: a space, and a note.
$pua = New-PdfBytes 'BT /F1 12 Tf <0041E0010042> Tj ET' -Plain -ToUnicode 'begincmap 1 beginbfrange <0000> <FFFF> <0000> endbfrange endcmap'
Assert-Equal $true ((Get-TextFromBytes -Bytes $pua -FileName 'p.pdf').Note -match 'Some characters in this PDF could not be read') 'PDF: private-use characters give a note'
