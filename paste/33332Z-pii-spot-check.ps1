<#
  Ion - DoD365 PII Reporter (33-332.Z; GENERATED from src/ by scripts/build_paste.py: do not edit this text)

  HOW TO RUN
    1. Optional: read this text first (the Open the script button saves it as a .txt file for Notepad).
    2. Open Windows PowerShell ISE and choose File > New.
    3. Paste ALL of this text (Ctrl+V) and press F5.
    4. When asked, paste a Teams link (... next to the team or channel name > Copy link) or a SharePoint
       site address (in Teams: the Shared tab > ... > Open in SharePoint). Sign in with your CAC.
       If you do not see a sign-in window, press Alt+Tab.
    5. Watch the two progress bars (the whole scan, and the file in work) and the line per folder.
       At the end, read the summary of hits per PII type. The results workbook is saved in
       Documents\Ion\<site>\<date>: a Summary tab, then one tab per PII type (worst first).
       Handle it as CUI: it shows where PII is.

  It scans the site's Documents library: every folder (in a team, one per standard channel) and the files at
  its top. One sign-in only: channel posts are not read.

  Built from these files; to change it, change them and rebuild with: python3 scripts/build_paste.py
    src/PiiDetectors.psm1, sha256 5758bbd9b581
    src/PiiMasking.psm1, sha256 dda7cdacfaf9
    src/TextExtract.psm1, sha256 7d3ab72e3627
    src/XlsxWriter.psm1, sha256 180fda2898e2
    src/Report.psm1, sha256 7ff23ecaac2d
    src/ScanMeter.psm1, sha256 9825bc0e2aa8
    src/ChannelScan.psm1, sha256 c1f5308abae5
    src/Find-ChannelPII.ps1, sha256 ccecf99228d2

  Read-only: it never changes, moves or deletes anything in Teams or SharePoint. Every value found is masked.
  Run target: Windows PowerShell 5.1 with PnP.PowerShell 1.12.0; no admin rights, no file to unblock.
  Network: only the one site given.
  Writes: only the results workbook under $OutputFolder (default Documents\Ion).
  Credentials: your own CAC sign-in in the PnP sign-in window; nothing is stored.

  Contents (line numbers of this file):
    line    48  Settings
    line    59  Module PiiDetectors.psm1
    line   523  Module PiiMasking.psm1
    line   709  Module TextExtract.psm1
    line   926  Module XlsxWriter.psm1
    line  1138  Module Report.psm1
    line  1504  Module ScanMeter.psm1
    line  1642  Module ChannelScan.psm1
    line  2051  Body of Find-ChannelPII.ps1
    line  2108  Clean-up
#>

# ==== Settings ==== (optional: edit the values between the quotes; empty = the script asks)
$SiteUrl = ''            # the team's Teams link, or its site https://usaf.dps.mil/teams/<team-site>; empty = asked
$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Ion')   # results: a folder per site, then per day
$MaxFileMB = 25          # files larger than this are reported as NOT SCANNED

# ISE runs pasted text in your session: keep its error setting so the Clean-up block at the end can put it back.
$previousErrorAction = $ErrorActionPreference
try {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ==== Module PiiDetectors.psm1 ====
<#
PiiDetectors.psm1 - finds text that looks like PII and returns where it is.

  Purpose:     One detector function per PII type (SSN, DoD ID, date of birth, payment card, passport,
               phone, personal e-mail, home address, labelled name). Get-PiiSpan runs them all, removes
               overlaps and returns each hit with its position and raw value.
  Inputs:      plain text (from TextExtract.psm1).
  Outputs:     hit objects WITH the raw value. Only PiiMasking.psm1 reads them; it masks every value
               before anything is shown or saved (the PAM, 2026-11-20: all values masked).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>


# =====================================================================================
# Settings (named values; each says where it comes from)
# =====================================================================================

# How far before a match to look for a label such as "SSN:" (characters). The author's choice,
# 2026-11-20: long enough for a label and a short phrase ("SSN of the member:"), short enough
# that a label in another table row does not count.
$script:KeywordWindowChars = 40

# A shorter window for labels that sit right next to their value ("DOB 01/02/1990", "cell 555-...").
$script:ShortKeywordWindowChars = 30

# A labelled name within this many characters of an identifier is raised to High. DoD privacy
# guidance treats a name together with an identifier (SSN, DoD ID, date of birth) as sensitive PII.
$script:NameToIdentifierChars = 80

# Overlap priority: when two hits cover the same characters, the type earlier in this list wins.
# Specific identifiers first, broad patterns (phone, address, name) last.
$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone',
    'Personal e-mail', 'Home address', 'Name')

# Types that make a nearby name sensitive (see $NameToIdentifierChars).
$script:IdentifierTypes = @('SSN', 'DoD ID', 'Date of birth', 'Passport', 'Payment card')

# Confidence labels as numbers, for sorting (High first).
$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# E-mail top-level labels that mark a government address (not reported as personal e-mail).
$script:GovernmentMailLabels = @('gov', 'mil')

# =====================================================================================
# Small helpers
# =====================================================================================

function Get-ConfidenceRank {
    # Returns 3, 2 or 1 for High, Medium or Low (0 for anything else), for sorting.
    param([string]$Confidence)
    if ($script:ConfidenceRank.ContainsKey($Confidence)) { return $script:ConfidenceRank[$Confidence] }
    return 0
}

function Test-KeywordBefore {
    <#
    .SYNOPSIS
      True when a label pattern (for example "SSN") appears just before a match.
    .PARAMETER Text
      The whole text.
    .PARAMETER Index
      Start position of the match in the text.
    .PARAMETER Pattern
      Regular expression of the labels (case is ignored).
    .PARAMETER Window
      How many characters before the match to search.
    .OUTPUTS
      [bool]
    #>
    param([string]$Text, [int]$Index, [string]$Pattern, [int]$Window = $script:KeywordWindowChars)
    $start = [Math]::Max(0, $Index - $Window)
    $before = $Text.Substring($start, $Index - $start)
    return [regex]::IsMatch($before, $Pattern, 'IgnoreCase')
}

function Test-SsnValid {
    <#
    .SYNOPSIS
      Applies the SSA issuing rules. The SSA never issues area 000, 666 or 900-999, group 00
      or serial 0000, so such a number is not an SSN.
    .PARAMETER Area
      The first 3 digits.
    .PARAMETER Group
      The middle 2 digits.
    .PARAMETER Serial
      The last 4 digits.
    .OUTPUTS
      [bool]
    #>
    param([string]$Area, [string]$Group, [string]$Serial)
    $areaNumber = [int]$Area
    if ($areaNumber -eq 0 -or $areaNumber -eq 666 -or $areaNumber -ge 900) { return $false }
    if ([int]$Group -eq 0) { return $false }
    if ([int]$Serial -eq 0) { return $false }
    return $true
}

function Test-Luhn {
    <#
    .SYNOPSIS
      Luhn (mod 10) check digit test, used by every payment card number (ISO/IEC 7812).
      It removes most random digit strings that only look like card numbers.
    .PARAMETER Digits
      Digits only, no spaces or dashes.
    .OUTPUTS
      [bool]
    #>
    param([string]$Digits)
    $sum = 0
    $doubleThis = $false
    # Walk from the right; double every second digit; subtract 9 from results over 9.
    for ($i = $Digits.Length - 1; $i -ge 0; $i--) {
        $digit = [int][string]$Digits[$i]
        if ($doubleThis) {
            $digit = $digit * 2
            if ($digit -gt 9) { $digit = $digit - 9 }
        }
        $sum += $digit
        $doubleThis = -not $doubleThis
    }
    return ($sum % 10 -eq 0)
}

function New-PiiHit {
    <#
    .SYNOPSIS
      Builds one hit object. It holds the raw value for overlap checks and masking;
      PiiMasking.psm1 removes the raw value before anything leaves the scan.
    .PARAMETER Detector
      The PII type, as named in $DetectorOrder.
    .PARAMETER Confidence
      High, Medium or Low.
    .PARAMETER Index
      Start position of the value in the text.
    .PARAMETER Value
      The raw value found.
    .OUTPUTS
      [pscustomobject] Detector, Confidence, Index, Length, Value.
    #>
    param([string]$Detector, [string]$Confidence, [int]$Index, [string]$Value)
    [pscustomobject]@{ Detector = $Detector; Confidence = $Confidence; Index = $Index; Length = $Value.Length; Value = $Value }
}

# =====================================================================================
# Detectors: one per PII type. Each returns zero or more hits (New-PiiHit).
# =====================================================================================

function Find-SsnHits {
    <#
    .SYNOPSIS
      Social Security numbers: 3-2-4 digits with the same separator (dash, space or none),
      valid by the SSA rules (Test-SsnValid).
    .DESCRIPTION
      High   = a label (SSN, SSAN, social security) is near.
      Medium = dashes or spaces, no label.
      Low    = 9 bare digits, no label (many false matches: check by hand).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = 'SSN|SSAN|social\s*security|soc\.?\s*sec'
    $pattern = '(?<![\d-])(\d{3})([- ]?)(\d{2})\2(\d{4})(?![\d-])'
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if (-not (Test-SsnValid $m.Groups[1].Value $m.Groups[3].Value $m.Groups[4].Value)) { continue }
        $confidence = 'Low'
        if ($m.Groups[2].Value -ne '') { $confidence = 'Medium' }
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
        New-PiiHit 'SSN' $confidence $m.Index $m.Value
    }
}

function Find-DodIdHits {
    <#
    .SYNOPSIS
      DoD ID numbers (EDIPI, 10 digits). Reported only with a label near (DoD ID, DODID, EDIPI, CAC),
      because 10 bare digits are far more often something else.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always High.
    #>
    param([string]$Text)
    $labels = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'
    foreach ($m in [regex]::Matches($Text, '(?<!\d)\d{10}(?!\d)')) {
        if (Test-KeywordBefore $Text $m.Index $labels) { New-PiiHit 'DoD ID' 'High' $m.Index $m.Value }
    }
}

function Find-CardHits {
    <#
    .SYNOPSIS
      Payment card numbers: 13-19 digits (ISO/IEC 7812 lengths), spaces or dashes allowed,
      that pass the Luhn check.
    .DESCRIPTION
      High with a label near (card, Visa, Mastercard, Amex, Discover, GTC, GPC), else Medium.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = 'card|visa|master\s*card|amex|american\s*express|discover|\bGTC\b|\bGPC\b'
    foreach ($m in [regex]::Matches($Text, '(?<![\d-])\d(?:[ -]?\d){12,18}(?![\d-])')) {
        $digits = $m.Value -replace '\D', ''
        if (-not (Test-Luhn $digits)) { continue }
        $confidence = 'Medium'
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
        New-PiiHit 'Payment card' $confidence $m.Index $m.Value
    }
}

function Find-DobHits {
    <#
    .SYNOPSIS
      Dates of birth: a date with a label near (DOB, date of birth, birth date, born).
      A date without a label is not reported (most dates are suspenses and events).
    .DESCRIPTION
      Date forms: 1/2/1990, 01-02-90, 1990-01-02, 2 Jan 1990, 02JAN1990, January 2, 1990.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always High.
    #>
    param([string]$Text)
    $labels = '\bDOB\b|date\s*of\s*birth|birth\s*date|birthdate|\bborn\b|\bD\.O\.B\.?'
    $month = '(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*'
    $numeric = '\d{1,2}[/-]\d{1,2}[/-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2}'
    $military = '\d{1,2}\s?' + $month + '\s?\d{2,4}'
    $written = $month + '\s+\d{1,2},?\s+\d{4}'
    $pattern = '\b(?:' + $numeric + '|' + $military + '|' + $written + ')\b'
    foreach ($m in [regex]::Matches($Text, $pattern, 'IgnoreCase')) {
        if (Test-KeywordBefore $Text $m.Index $labels $script:ShortKeywordWindowChars) { New-PiiHit 'Date of birth' 'High' $m.Index $m.Value }
    }
}

function Find-PassportHits {
    <#
    .SYNOPSIS
      Passport numbers: 6-9 capital letters or digits (at least 5 digits) with "passport" near.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always Medium.
    #>
    param([string]$Text)
    foreach ($m in [regex]::Matches($Text, '\b[A-Z0-9]{6,9}\b')) {
        if (($m.Value -replace '\D', '').Length -lt 5) { continue }
        if (Test-KeywordBefore $Text $m.Index 'passport' $script:ShortKeywordWindowChars) { New-PiiHit 'Passport' 'Medium' $m.Index $m.Value }
    }
}

function Find-PhoneHits {
    <#
    .SYNOPSIS
      US phone numbers (10 digits, optional +1). Medium with a personal label near
      (cell, mobile, home, personal, my number), else Low (office numbers are common).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $labels = '\bcell\b|mobile|\bhome\b|personal|\bmy\s+(?:number|phone)'
    $pattern = '(?<![\d-])(?:\+?1[ .-]?)?(?:\(\d{3}\)\s?|\d{3}[ .-])\d{3}[ .-]\d{4}(?![\d-])'
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        $confidence = 'Low'
        if (Test-KeywordBefore $Text $m.Index $labels $script:ShortKeywordWindowChars) { $confidence = 'Medium' }
        New-PiiHit 'Phone' $confidence $m.Index $m.Value
    }
}

function Find-PersonalEmailHits {
    <#
    .SYNOPSIS
      E-mail addresses outside government domains (last label not in $GovernmentMailLabels).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), always Medium.
    #>
    param([string]$Text)
    # Local part: any characters but spaces, delimiters and path slashes (accents and apostrophes included),
    # starting after a space, a delimiter, a slash, an '@' or the start of the text. Domain: any such
    # characters too (non-ASCII domains), ending in a top-level label of 2 or more letters. No further '@' may
    # follow before the next space or delimiter: in a run of two addresses glued by '@', the first part is not
    # taken for an address whose 'domain' is the second address's local part. Any other trailing character
    # (? ! quotes ... and so on) still ends the address normally. The top-level label is atomic ((?>...)): without it,
    # a very long run of letters with no space makes the end check retry at every letter (seconds per 100,000 letters).
    # The domain uses the internet's own limits (RFC 1035: a label is at most 63 characters; at most 10 labels here;
    # a top-level label of 2 to 24 letters), so a long dotted run without spaces is tried a fixed number of ways,
    # not once per label: the time grows with the text, not with its square. The address ends where its domain ends:
    # no letter or digit right after the top-level label, and no letter or digit right after a final period (so an
    # address that ends a sentence, then a quote, a dash or a star, is still found). The no-further-'@' look-ahead
    # stops at the next '@', so a long run of glued addresses is read once, not again for every address in it.
    $notInAddress = '\s@<>()\[\],;:"/\\'
    $domain = '((?:[^' + $notInAddress + '.]{1,63}\.){1,10}((?>\p{L}{2,24})))'
    $endsAddress = '(?![\p{L}\p{N}])(?!\.[\p{L}\p{N}])(?![^\s<>()\[\],;:"/\\@]*@)'
    $pattern = '(?<=^|[' + $notInAddress + '])[^' + $notInAddress + ']+@' + $domain + $endsAddress
    foreach ($m in [regex]::Matches($Text, $pattern)) {
        if ($script:GovernmentMailLabels -contains $m.Groups[2].Value.ToLower()) { continue }
        New-PiiHit 'Personal e-mail' 'Medium' $m.Index $m.Value
    }
}

function Find-HomeAddressHits {
    <#
    .SYNOPSIS
      US street addresses: house number, 1-4 capitalised words and a street type (St, Ave, Rd ...),
      with an optional unit, city, state and ZIP code. Added at the PAM's request, 2026-11-20.
    .DESCRIPTION
      High   = a label near (home address, residence, mailing address, lives at, home).
      Medium = ends with a 2-letter state and a ZIP code.
      Low    = street only (it can be a work address: check by hand).
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    $streetTypes = 'Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Court|Ct|Boulevard|Blvd|Way|Circle|Cir|' +
                   'Place|Pl|Terrace|Ter|Parkway|Pkwy|Highway|Hwy|Trail|Trl'
    $street = '\b\d{1,6}\s+(?:[NSEW]\.?\s+)?(?:[A-Z][A-Za-z]+\s+){1,4}(?:' + $streetTypes + ')\b\.?'
    $unit = '(?:(?:Apt|Unit|#)\s*\w+,?\s*)?'
    $cityStateZip = '(?:,?\s*' + $unit + '[A-Z][A-Za-z]+(?:\s[A-Z][A-Za-z]+)?,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?)?'
    $labels = 'home\s*address|residence|residential|mailing\s*address|lives\s+at|\bhome\b'
    foreach ($m in [regex]::Matches($Text, $street + $cityStateZip)) {
        $confidence = 'Low'
        if ($m.Value -match '\b[A-Z]{2}\s+\d{5}') { $confidence = 'Medium' }
        if (Test-KeywordBefore $Text $m.Index $labels) { $confidence = 'High' }
        New-PiiHit 'Home address' $confidence $m.Index $m.Value
    }
}

function Find-NameHits {
    <#
    .SYNOPSIS
      Names after a label (Name:, Full name:, Member:, Employee:, Spouse:, Dependent:, NOK:,
      Next of kin:, POC:). Added at the PAM's request, 2026-11-20.
    .DESCRIPTION
      Low by itself: a name alone is PII but rarely sensitive. Get-PiiSpan raises it to High when an
      identifier is near ($NameToIdentifierChars). A name without a label is not found: that would
      need a list of real names. Military "LAST, FIRST MI" form is found.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit) for the name only, not its label.
    #>
    param([string]$Text)
    $label = '(?:full\s+name|member\s+name|name|member|employee|spouse|dependent|next\s+of\s+kin|NOK|POC)\s*[:=-]\s*'
    # Labels of other data often follow a name on the same line; a name stops before them.
    $stopWords = '(?!(?:SSN|SSAN|DOB|DoD|EDIPI|CAC|Phone|Cell|Mobile|Home|Address|Email|E-mail|Rank|Grade|Unit|Date|Born)\b)'
    $name = "[A-Z][A-Za-z'\-]+(?:,?\s+$stopWords[A-Z](?:[A-Za-z'\-]+|\.))+"
    foreach ($m in [regex]::Matches($Text, '(?i:' + $label + ')(' + $name + ')')) {
        New-PiiHit 'Name' 'Low' $m.Groups[1].Index $m.Groups[1].Value
    }
}

# =====================================================================================
# All detectors together
# =====================================================================================

function Get-AllHits {
    <#
    .SYNOPSIS
      Runs every detector on the text and returns all hits, overlaps included.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit).
    #>
    param([string]$Text)
    Find-SsnHits $Text
    Find-DodIdHits $Text
    Find-CardHits $Text
    Find-DobHits $Text
    Find-PassportHits $Text
    Find-PhoneHits $Text
    Find-PersonalEmailHits $Text
    Find-HomeAddressHits $Text
    Find-NameHits $Text
}

function Remove-OverlappingHits {
    <#
    .SYNOPSIS
      Keeps one hit per stretch of text: the type earlier in $DetectorOrder wins, then the higher
      confidence. Example: the 9 digits inside a 16-digit card number are not also an SSN.
    .DESCRIPTION
      The kept hits never overlap, so kept in order of their start they also end in order. A new hit overlaps a
      kept one only if it overlaps the kept hit that starts last before the new hit ends: one binary search per
      hit instead of a check against every kept hit. (Each kept hit is inserted in place, a fast memory move.)
    .PARAMETER Hits
      All hits from Get-AllHits.
    .OUTPUTS
      The kept hits (a list, in priority order).
    #>
    param([object[]]$Hits)
    $sorted = $Hits | Where-Object { $_ } | Sort-Object `
        @{ Expression = { [array]::IndexOf($script:DetectorOrder, $_.Detector) } },
        @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true },
        Index
    $kept = New-Object System.Collections.Generic.List[object]
    $starts = New-Object System.Collections.Generic.List[int]   # starts of the kept hits, ascending
    $ends = New-Object System.Collections.Generic.List[int]     # their ends, in the same order (also ascending)
    foreach ($hit in $sorted) {
        $end = $hit.Index + $hit.Length
        # The first kept hit that starts at or after this hit's end (BinarySearch gives its place as ~n if absent).
        $after = $starts.BinarySearch($end)
        if ($after -lt 0) { $after = -bnot $after }
        while ($after -gt 0 -and $starts[$after - 1] -ge $end) { $after-- }
        if ($after -gt 0 -and $ends[$after - 1] -gt $hit.Index) { continue }   # overlaps the kept hit before it
        $starts.Insert($after, $hit.Index)
        $ends.Insert($after, $end)
        $kept.Add($hit)
    }
    return , $kept
}

function Set-NameNearIdentifierHigh {
    <#
    .SYNOPSIS
      Raises a labelled name to High when an identifier hit is within $NameToIdentifierChars.
    .PARAMETER Hits
      The kept hits (changed in place).
    .OUTPUTS
      None.
    #>
    param([object[]]$Hits)
    foreach ($nameHit in @($Hits | Where-Object { $_.Detector -eq 'Name' })) {
        foreach ($other in $Hits) {
            $near = [Math]::Abs($other.Index - $nameHit.Index) -le $script:NameToIdentifierChars
            if ($near -and $script:IdentifierTypes -contains $other.Detector) { $nameHit.Confidence = 'High'; break }
        }
    }
}

function Get-PiiSpan {
    <#
    .SYNOPSIS
      The one entry point of this module: every PII hit in a text, without overlaps, in text order.
      Hits hold raw values: pass them only to PiiMasking.psm1.
    .PARAMETER Text
      The text to search.
    .OUTPUTS
      Hit objects (New-PiiHit), sorted by position.
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    # Underscores join words in file names ("CAC_1234..."), which defeats word-boundary checks.
    # Seen as spaces by the detectors only; the length is the same, so positions still fit the text.
    $searchText = $Text.Replace('_', ' ')
    $kept = Remove-OverlappingHits @(Get-AllHits $searchText)
    Set-NameNearIdentifierHigh $kept.ToArray()
    foreach ($hit in $kept) { $hit.Value = $Text.Substring($hit.Index, $hit.Length) }  # the value as written
    return @($kept | Sort-Object Index)
}

# ==== Module PiiMasking.psm1 ====
<#
PiiMasking.psm1 - turns raw PII hits into masked report data. Nothing unmasked leaves this module.

  Purpose:     Find-PiiInText returns, for one text, each hit's type, confidence, masked value and a
               masked snippet of the text around it. Protect-PiiText masks PII inside a short text that
               the report shows as it is (file and folder names, error messages).
               The PAM, 2026-11-20: every value is masked on screen and in the report files.
  Inputs:      plain text; hits from PiiDetectors.psm1 (Get-PiiSpan).
  Outputs:     masked hit objects and masked strings.
  Depends on:  PiiDetectors.psm1 (Get-PiiSpan), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
#>


# =====================================================================================
# Settings
# =====================================================================================

# Characters of text kept on each side of a hit in the report snippet. The author's choice,
# 2026-11-20: enough to see the label ("SSN:", "DOB") without copying a whole paragraph.
$script:SnippetRadiusChars = 30

# How many trailing characters a masked value keeps, so a reviewer can find it in the file. The author's
# choice, 2026-11-20, after common practice: receipts and DoD forms show at most the last 4 digits of an
# SSN or card number. Passports keep 3, as their numbers are shorter.
$script:KeepLastDigits = 4
$script:KeepLastPassportChars = 3

# An e-mail local part: everything before "@domain.tld" that is not a space, a delimiter or a path slash.
# Broad on purpose (apostrophes, accents, non-ASCII domains), so no part of a name in an address stays
# visible; slashes are excluded so the folders before an address in a path stay readable.
# It starts only at the start of a word (after a space, a delimiter, a slash or an '@'), so a long run of text is read
# once, not once per character.
$script:EmailLocalPartPattern = '(?<=^|[\s@<>()\[\],;:"/\\])[^\s@<>()\[\],;:"/\\]+(?=@[^\s@<>()\[\],;:"/\\.]+\.)'

# =====================================================================================
# Masking one value
# =====================================================================================

function Get-LastCharacters {
    # Returns the last $Count characters of a string (all of it when it is shorter).
    param([string]$Value, [int]$Count)
    return $Value.Substring([Math]::Max(0, $Value.Length - $Count))
}

function Get-MaskedValue {
    <#
    .SYNOPSIS
      Masks one found value so the report never holds it in full.
    .DESCRIPTION
      SSN    -> ***-**-1234        DoD ID   -> ******7890        Card    -> ************1111
      Phone  -> ***-***-1234       Passport -> *****567          E-mail  -> ***@domain
      DOB    -> [date masked]      Address  -> [address masked]  Name    -> initials (P*** S***)
    .PARAMETER Detector
      The PII type (as named in PiiDetectors.psm1).
    .PARAMETER Value
      The raw value.
    .OUTPUTS
      [string] the masked value.
    #>
    param([string]$Detector, [string]$Value)
    $digits = $Value -replace '\D', ''
    $lastDigits = Get-LastCharacters $digits $script:KeepLastDigits
    switch ($Detector) {
        'SSN' { return '***-**-' + $lastDigits }
        'DoD ID' { return '******' + $lastDigits }
        'Payment card' { return ('*' * ($digits.Length - $lastDigits.Length)) + $lastDigits }
        'Phone' { return '***-***-' + $lastDigits }
        'Personal e-mail' { return '***@' + (($Value -split '@')[-1] -replace '\d', '#') }   # domain kept, its digits hidden
        'Passport' {
            $tail = Get-LastCharacters $Value $script:KeepLastPassportChars
            return ('*' * ($Value.Length - $tail.Length)) + $tail
        }
        'Date of birth' { return '[date masked]' }
        'Home address' { return '[address masked]' }
        'Name' { return ((($Value -split '[\s,]+') | Where-Object { $_ } | ForEach-Object { $_.Substring(0, 1) + '***' }) -join ' ') }
        default { return '[masked]' }
    }
}

# =====================================================================================
# Masking a whole text, then cutting snippets from it
# =====================================================================================

function Get-MaskedText {
    <#
    .SYNOPSIS
      Masks a whole text and keeps its length: every digit becomes '#', every e-mail local part
      becomes '*', and every hit (names and addresses included) becomes '*'. Because the length
      stays the same, hit positions from the raw text still point to the right place.
    .PARAMETER Text
      The raw text.
    .PARAMETER Hits
      The hits found in it (Get-PiiSpan).
    .OUTPUTS
      [string] the masked text, same length as the input.
    #>
    param([string]$Text, [object[]]$Hits)
    $sameLengthStars = [System.Text.RegularExpressions.MatchEvaluator] { param($m) '*' * $m.Length }
    $masked = [regex]::Replace($Text, $script:EmailLocalPartPattern, $sameLengthStars) -replace '\d', '#'
    $chars = $masked.ToCharArray()
    foreach ($hit in $Hits) {
        for ($i = $hit.Index; $i -lt $hit.Index + $hit.Length; $i++) {
            if (-not [char]::IsWhiteSpace($chars[$i])) { $chars[$i] = '*' }
        }
    }
    return New-Object string (, $chars)
}

function Get-MaskedSnippet {
    <#
    .SYNOPSIS
      A short piece of the MASKED text around one hit. It is cut from text that is already masked,
      so a cut through an address or number cannot expose part of it.
    .PARAMETER MaskedText
      Output of Get-MaskedText.
    .PARAMETER Index
      Start position of the hit.
    .PARAMETER Length
      Length of the hit.
    .OUTPUTS
      [string] the snippet, white space folded to single spaces.
    #>
    param([string]$MaskedText, [int]$Index, [int]$Length)
    $start = [Math]::Max(0, $Index - $script:SnippetRadiusChars)
    $end = [Math]::Min($MaskedText.Length, $Index + $Length + $script:SnippetRadiusChars)
    return ($MaskedText.Substring($start, $end - $start) -replace '\s+', ' ').Trim()
}

# =====================================================================================
# Public functions
# =====================================================================================

function Find-PiiInText {
    <#
    .SYNOPSIS
      Finds PII in one text and returns masked results only.
    .PARAMETER Text
      Plain text from a file or a post.
    .OUTPUTS
      [pscustomobject] Detector, Confidence, MaskedValue, Context (no raw values).
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return }
    $hits = @(Get-PiiSpan $Text)
    if ($hits.Count -eq 0) { return }
    $maskedText = Get-MaskedText $Text $hits
    foreach ($hit in $hits) {
        [pscustomobject]@{
            Detector    = $hit.Detector
            Confidence  = $hit.Confidence
            MaskedValue = Get-MaskedValue $hit.Detector $hit.Value
            Context     = Get-MaskedSnippet $maskedText $hit.Index $hit.Length
        }
    }
}

function Protect-PiiText {
    <#
    .SYNOPSIS
      Masks a short text that the report shows as it is: file, folder and channel names, and error
      messages. Strict on purpose, because a detector can miss a pattern in a name: each found value
      becomes its masked value, then EVERY digit becomes '#' and every e-mail name becomes '***'.
      Example: "Roster_123-45-6789.docx" -> "Roster_***-**-####.docx"; "FY27 Plan.docx" -> "FY## Plan.docx".
    .PARAMETER Text
      The name or message.
    .OUTPUTS
      [string]
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $result = $Text
    # Replace found values from the end, so the positions of earlier ones stay valid.
    foreach ($hit in (@(Get-PiiSpan $Text) | Sort-Object Index -Descending)) {
        $masked = Get-MaskedValue $hit.Detector $hit.Value
        $result = $result.Substring(0, $hit.Index) + $masked + $result.Substring($hit.Index + $hit.Length)
    }
    $result = [regex]::Replace($result, $script:EmailLocalPartPattern, '***')
    return ($result -replace '\d', '#')
}

# ==== Module TextExtract.psm1 ====
<#
TextExtract.psm1 - turns file bytes into plain text, in memory.

  Purpose:     Plain text from Word, Excel and PowerPoint files (.docx/.xlsx/.pptx and macro versions),
               text files and HTML. Other types (PDF, images, legacy .doc/.xls/.ppt, scripts and
               anything else) are reported as NOT SCANNED, so a person checks them by hand.
  Inputs:      file bytes and the file name (for its extension).
  Outputs:     an object with Status ('OK' or 'NOT SCANNED'), Reason and Text.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7); .NET System.IO.Compression only.
  Network:     none.
  Writes:      none. Bytes stay in memory; nothing is saved to disk.
  Credentials: none.
#>

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
    $text = $Xml -replace '<w:tab/>|<w:br/>|<a:br/>', ' '
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

# ==== Module XlsxWriter.psm1 ====
<#
XlsxWriter.psm1 - writes a small Excel workbook (.xlsx) with .NET only: no Excel, no module.

  Purpose:     The results file of Ion: one workbook with tabs (the PAM, 2026-11-20: a CSV cannot have tabs).
               Each tab has a bold header row; a list tab also has the header frozen and filter buttons on.
  Inputs:      the tabs to write (New-XlsxSheet), each with its columns and rows of cells (text or whole numbers).
  Outputs:     one .xlsx file (Office Open XML, ECMA-376: a zip of XML parts).
  Depends on:  .NET System.IO.Compression (loaded here).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the one .xlsx file the caller names; it never overwrites a file that exists.
  Credentials: none.

  Safety:      every text cell is an inline string, never a formula, so text from a scanned file cannot run in Excel.
#>

Add-Type -AssemblyName System.IO.Compression

# Excel limits: 31 characters in a tab name, 32,767 in a cell (ECMA-376 and Excel's specifications and limits).
$script:MaxSheetName = 31
$script:MaxCellText = 32767

# XML namespaces of the parts.
$script:MainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Two cell styles: 0 = normal, 1 = bold (the header row and the Summary labels).
$script:StylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>'

function New-XlsxSheet {
    <#
    .SYNOPSIS
      One tab of the workbook.
    .PARAMETER Name
      The tab name (cut to 31 characters; characters Excel refuses are replaced).
    .PARAMETER Columns
      The header row.
    .PARAMETER Rows
      The rows: each an array of cells, a cell being text or a whole number.
    .PARAMETER Widths
      Column widths in characters, one per column.
    .PARAMETER List
      Freeze the header row and turn the filter buttons on (a list of hits); off for the Summary tab.
    .OUTPUTS
      [pscustomobject]
    #>
    param([string]$Name, [string[]]$Columns, [object[]]$Rows, [int[]]$Widths, [switch]$List)
    [pscustomobject]@{ Name = $Name; Columns = $Columns; Rows = @($Rows); Widths = $Widths; List = [bool]$List }
}

function ConvertTo-XmlText {
    # Text for an XML part: characters XML does not allow are dropped, the rest is escaped, and it is cut to Excel's cell limit.
    param([AllowEmptyString()][string]$Text)
    $clean = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', ''
    if ($clean.Length -gt $script:MaxCellText) { $clean = $clean.Substring(0, $script:MaxCellText) }
    return [System.Security.SecurityElement]::Escape($clean)
}

function Get-ColumnLetter {
    # The column letters of a 1-based column number: 1 = A, 27 = AA.
    param([int]$Number)
    $letters = ''
    while ($Number -gt 0) { $Number--; $letters = [char](65 + $Number % 26) + $letters; $Number = [Math]::Floor($Number / 26) }
    return $letters
}

function ConvertTo-CellXml {
    # One cell: a whole number as a number, anything else as an inline string (never a formula).
    param([string]$Ref, $Value, [int]$Style)
    $styleAttr = ''
    if ($Style -gt 0) { $styleAttr = ' s="' + $Style + '"' }
    if ($Value -is [int] -or $Value -is [long]) { return '<c r="{0}"{1}><v>{2}</v></c>' -f $Ref, $styleAttr, $Value }
    return '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Ref, $styleAttr, (ConvertTo-XmlText ([string]$Value))
}

function ConvertTo-SheetXml {
    <#
    .SYNOPSIS
      The worksheet XML of one tab.
    .PARAMETER Sheet
      Output of New-XlsxSheet.
    .OUTPUTS
      [string]
    #>
    param($Sheet)
    $xml = New-Object System.Text.StringBuilder
    [void]$xml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + $script:MainNs + '">')
    if ($Sheet.List) {
        [void]$xml.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    }
    if ($Sheet.Widths) {
        [void]$xml.Append('<cols>')
        for ($i = 0; $i -lt $Sheet.Widths.Count; $i++) {
            [void]$xml.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $Sheet.Widths[$i]))
        }
        [void]$xml.Append('</cols>')
    }
    [void]$xml.Append('<sheetData><row r="1">')
    for ($i = 0; $i -lt $Sheet.Columns.Count; $i++) { [void]$xml.Append((ConvertTo-CellXml ((Get-ColumnLetter ($i + 1)) + '1') $Sheet.Columns[$i] 1)) }
    [void]$xml.Append('</row>')
    $rowNumber = 1
    foreach ($row in $Sheet.Rows) {
        $rowNumber++
        [void]$xml.Append('<row r="' + $rowNumber + '">')
        $cells = @($row)
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $style = 0
            if (-not $Sheet.List -and $i -eq 0) { $style = 1 }   # Summary tab: the label column is bold
            [void]$xml.Append((ConvertTo-CellXml ((Get-ColumnLetter ($i + 1)) + $rowNumber) $cells[$i] $style))
        }
        [void]$xml.Append('</row>')
    }
    [void]$xml.Append('</sheetData>')
    if ($Sheet.List -and $Sheet.Columns.Count -gt 0) {
        [void]$xml.Append(('<autoFilter ref="A1:{0}{1}"/>' -f (Get-ColumnLetter $Sheet.Columns.Count), [Math]::Max(2, $rowNumber)))
    }
    [void]$xml.Append('</worksheet>')
    return $xml.ToString()
}

function Get-SafeSheetNames {
    <#
    .SYNOPSIS
      Tab names Excel accepts: no : \ / ? * [ ], at most 31 characters, none twice (a number is added).
    .PARAMETER Names
      The wanted names, in order.
    .OUTPUTS
      [string[]] the names to use, in the same order.
    #>
    param([string[]]$Names)
    $used = @{}
    foreach ($name in $Names) {
        $clean = ($name -replace '[:\\/?*\[\]]', '-').Trim()
        if (-not $clean) { $clean = 'Sheet' }
        if ($clean.Length -gt $script:MaxSheetName) { $clean = $clean.Substring(0, $script:MaxSheetName) }
        $candidate = $clean
        $n = 2
        while ($used.ContainsKey($candidate.ToLower())) {
            $suffix = ' (' + $n + ')'
            $candidate = $clean.Substring(0, [Math]::Min($clean.Length, $script:MaxSheetName - $suffix.Length)) + $suffix
            $n++
        }
        $used[$candidate.ToLower()] = $true
        $candidate
    }
}

function Add-ZipText {
    # Adds one UTF-8 text part (no byte order mark) to the zip.
    param($Zip, [string]$Name, [string]$Text)
    $writer = New-Object System.IO.StreamWriter($Zip.CreateEntry($Name).Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Export-Xlsx {
    <#
    .SYNOPSIS
      Writes the tabs to a new .xlsx file. Throws if the file exists (the caller picks a new name).
    .PARAMETER Path
      The new file.
    .PARAMETER Sheets
      The tabs, in order (New-XlsxSheet).
    #>
    param([string]$Path, [object[]]$Sheets)
    $names = @(Get-SafeSheetNames @($Sheets | ForEach-Object { $_.Name }))
    $types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $book = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + $script:MainNs + '" xmlns:r="' + $script:RelNs + '"><sheets>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">'
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $types += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        $book += '<sheet name="' + (ConvertTo-XmlText $names[$i - 1]) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'
        $rels += '<Relationship Id="rId' + $i + '" Type="' + $script:RelNs + '/worksheet" Target="worksheets/sheet' + $i + '.xml"/>'
    }
    $rels += '<Relationship Id="rId' + ($Sheets.Count + 1) + '" Type="' + $script:RelNs + '/styles" Target="styles.xml"/></Relationships>'
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::CreateNew)   # CreateNew: never overwrite
    $zip = $null
    $written = $false
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        Add-ZipText $zip '[Content_Types].xml' ($types + '</Types>')
        Add-ZipText $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $script:PackageRelNs + '">' +
            '<Relationship Id="rId1" Type="' + $script:RelNs + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        Add-ZipText $zip 'xl/workbook.xml' ($book + '</sheets></workbook>')
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' $rels
        Add-ZipText $zip 'xl/styles.xml' $script:StylesXml
        for ($i = 1; $i -le $Sheets.Count; $i++) { Add-ZipText $zip ('xl/worksheets/sheet' + $i + '.xml') (ConvertTo-SheetXml $Sheets[$i - 1]) }
        $written = $true
    }
    finally {
        if ($zip) { $zip.Dispose() }
        $stream.Dispose()
        # A half-written workbook would not open: remove it, so a rerun can use the name.
        if (-not $written) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
    }
}

# ==== Module Report.psm1 ====
<#
Report.psm1 - the scan results for a person: a summary on screen and one Excel workbook with tabs.

  Purpose:     Holds the report rows, sorts them worst PII type first, shows the summary (totals and hits per type)
               and saves the workbook: Summary tab first, one tab per type, Low confidence, Not scanned
               (the PAM, 2026-11-20: tabs to cut noise; no .txt or .csv).
  Inputs:      report rows (New-ReportRow) from ChannelScan.psm1; values in them are already masked.
  Outputs:     the summary on screen; <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx.
  Depends on:  PiiMasking.psm1 (Protect-PiiText), XlsxWriter.psm1, imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      the workbook and its site and day folders under the output folder (default Documents\Ion).
               A new file each run; an existing result is never overwritten.
  Credentials: none.
#>


# =====================================================================================
# Settings
# =====================================================================================

# Confidence labels as numbers, for sorting (High first). Rows with other labels are notes
# (NOT SCANNED, BLOCKED, INCOMPLETE).
$script:Rank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

# The tool's name on screen and in the workbook (the PAM, 2026-11-20).
$script:ToolTitle = 'Ion - DoD365 PII Reporter'

# PII types from most to least damaging, for sorting and for the order of the tabs. AFI 33-332 gives no ranking
# (A3.1 lists only the factors for judging harm), so this is the PAM's order (2026-11-20): SSN and date of birth
# first, then identity documents and accounts, then contact details.
$script:SeverityOrder = @('SSN', 'Date of birth', 'Passport', 'Payment card', 'DoD ID', 'Home address', 'Name', 'Phone', 'Personal e-mail')

# How many files the screen summary lists (the workbook lists all).
$script:TopItems = 10

# Line width for the text tables. Without a console window, Out-String has no width and prints
# empty tables, so a width is always given. 220 fits the widest table on a landscape printout.
$script:TableWidth = 220

# What to do with PII found where it should not be: the breach steps of the rule source, AFI 33-332 para 3.1.2 as
# changed by AFI33-332_DAFGM2026-01 (7 July 2026; edition: docs/DECISIONS.md).
# The wording is the same on the site page (the PAM, 2026-11-20).
$script:BreachAdvice = 'PII where it should not be is a potential or actual breach; report it to your servicing Privacy ' +
                       'Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, page 1, ' +
                       'second paragraph).'

# =====================================================================================
# Rows and summary
# =====================================================================================

function New-ReportRow {
    <#
    .SYNOPSIS
      One report row: a possible-PII hit, or a note (NOT SCANNED, BLOCKED, INCOMPLETE) about an item.
    .PARAMETER Source
      File, File name or Scan.
    .PARAMETER Channel
      The channel (folder) name.
    .PARAMETER Item
      The file path inside the folder (masked by the caller).
    .PARAMETER Link
      A link to the item, or empty.
    .PARAMETER Modified
      When the item last changed, or empty.
    .PARAMETER Detector
      The PII type, or the note word.
    .PARAMETER Confidence
      High, Medium, Low, or empty for a note.
    .PARAMETER MaskedValue
      The masked value, or empty.
    .PARAMETER Context
      The masked snippet, or the reason for a note.
    .OUTPUTS
      [pscustomobject] Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link.
    #>
    param(
        [string]$Source, [string]$Channel, [string]$Item, [string]$Link, [string]$Modified,
        [string]$Detector, [string]$Confidence, [string]$MaskedValue, [string]$Context
    )
    [pscustomobject]@{
        Source = $Source; Channel = $Channel; Item = $Item; Detector = $Detector; Confidence = $Confidence
        MaskedValue = $MaskedValue; Context = $Context; Modified = $Modified; Link = $Link
    }
}

function Add-IncompleteRow {
    <#
    .SYNOPSIS
      Adds one INCOMPLETE row when the scan did not reach its end and no INCOMPLETE row is there yet.
      Called from the script's 'finally' block, which also runs after Ctrl+C (when 'catch' does not).
    .PARAMETER Rows
      The list of report rows (changed in place).
    .PARAMETER Finished
      True when every step ran.
    .OUTPUTS
      None.
    #>
    param($Rows, [bool]$Finished)
    if ($Finished) { return }
    if (@($Rows | Where-Object { $_.Detector -eq 'INCOMPLETE' }).Count -gt 0) { return }
    $Rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' 'Stopped before the end (cancelled?)'))
}

function Get-SeverityIndex {
    # Where a PII type stands in $SeverityOrder (0 = worst); a type not in the list comes last.
    param([string]$Type)
    $index = [Array]::IndexOf($script:SeverityOrder, $Type)
    if ($index -lt 0) { return $script:SeverityOrder.Count }
    return $index
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
      One line per item with possible PII: its worst type, highest confidence, number of hits and types found.
    .PARAMETER Rows
      All report rows.
    .OUTPUTS
      [pscustomobject] Worst, Highest, Hits, Channel, Source, Item, Detectors; worst type first, then highest confidence.
    #>
    param([object[]]$Rows)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $summary = foreach ($group in ($hits | Group-Object Channel, Item)) {
        $first = $group.Group[0]
        $worst = ($group.Group | Sort-Object { Get-SeverityIndex $_.Detector }, { -$script:Rank[$_.Confidence] } | Select-Object -First 1)
        $best = ($group.Group | Sort-Object { $script:Rank[$_.Confidence] } -Descending | Select-Object -First 1).Confidence
        [pscustomobject]@{
            Worst = $worst.Detector; Highest = $best; Hits = $group.Count; Channel = $first.Channel; Source = $first.Source; Item = $first.Item
            Detectors = ((@($group.Group | Select-Object -ExpandProperty Detector -Unique) | Sort-Object { Get-SeverityIndex $_ }) -join ', ')
        }
    }
    return @($summary | Sort-Object @{ Expression = { Get-SeverityIndex $_.Worst } }, @{ Expression = { $script:Rank[$_.Highest] }; Descending = $true },
        @{ Expression = 'Hits'; Descending = $true })
}

function ConvertTo-TableText {
    <#
    .SYNOPSIS
      Formats objects as a text table at a fixed width (see $TableWidth).
    .PARAMETER InputRows
      The objects to show.
    .PARAMETER Columns
      The property names to show, in order.
    .OUTPUTS
      [string]
    #>
    param([object[]]$InputRows, [string[]]$Columns)
    return ($InputRows | Format-Table -Property $Columns -AutoSize -Wrap | Out-String -Width $script:TableWidth).TrimEnd()
}

function Get-TypeCounts {
    <#
    .SYNOPSIS
      The hits per PII type, worst type first, split by confidence.
    .PARAMETER Hits
      The hit rows.
    .OUTPUTS
      [pscustomobject] Type, High, Medium, Low, Total.
    #>
    param([object[]]$Hits)
    foreach ($group in ($Hits | Group-Object Detector | Sort-Object { Get-SeverityIndex $_.Name })) {
        $count = @{ High = 0; Medium = 0; Low = 0 }
        foreach ($hit in $group.Group) { $count[$hit.Confidence]++ }
        [pscustomobject]@{ Type = $group.Name; High = $count.High; Medium = $count.Medium; Low = $count.Low; Total = $group.Count }
    }
}

function Get-ScanMetrics {
    <#
    .SYNOPSIS
      The scan's totals, as label and value pairs (the screen summary and the Summary tab show the same).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .OUTPUTS
      [System.Collections.Specialized.OrderedDictionary]
    #>
    param([object[]]$Rows, [int]$ItemsScanned)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $metrics = [ordered]@{}
    $metrics['Files read'] = $ItemsScanned
    $metrics['Files with possible PII'] = @(Get-ReportSummary $Rows).Count
    $metrics['Possible PII hits'] = $hits.Count
    $metrics['High / Medium / Low'] = '{0} / {1} / {2}' -f @($hits | Where-Object { $_.Confidence -eq 'High' }).Count,
        @($hits | Where-Object { $_.Confidence -eq 'Medium' }).Count, @($hits | Where-Object { $_.Confidence -eq 'Low' }).Count
    $metrics['Not scanned / incomplete'] = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) }).Count
    return $metrics
}

function Get-ScanSummaryText {
    <#
    .SYNOPSIS
      The summary for the screen: the scan's details and totals, possible PII per type (worst first), and the
      files with the most serious hits. Every hit is listed in the workbook, not here.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string[]]
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(''); $lines.Add('==== ' + $script:ToolTitle + ' ====')
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) {
        $lines.Add(('{0,-26} {1}' -f ($pair.Key + ':'), $pair.Value))
    }
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    if ($hits.Count -gt 0) {
        $lines.Add(''); $lines.Add('POSSIBLE PII BY TYPE (worst first):')
        $lines.Add((ConvertTo-TableText @(Get-TypeCounts $hits) @('Type', 'High', 'Medium', 'Low', 'Total')))
        $top = @(Get-ReportSummary $Rows | Select-Object -First $script:TopItems |
            Select-Object Worst, Highest, Hits, @{ Name = $FolderLabel; Expression = { $_.Channel } }, Item)
        $lines.Add(''); $lines.Add(('FILES WITH THE MOST SERIOUS HITS (first {0}; all are in the workbook):' -f $script:TopItems))
        $lines.Add((ConvertTo-TableText $top @('Worst', 'Highest', 'Hits', $FolderLabel, 'Item')))
    }
    else { $lines.Add(''); $lines.Add('No possible PII found in the files that were read.') }
    $lines.Add('')
    $lines.Add('Confidence is a pattern guess, not proof: open each item and check it before you act.')
    $lines.Add($script:BreachAdvice)
    return $lines.ToArray()
}

# =====================================================================================
# The workbook
# =====================================================================================

function Get-ResultPath {
    <#
    .SYNOPSIS
      A new workbook path: <OutputFolder>\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx, with (2), (3) ... added
      if the name exists, so a result is never overwritten. Creates the folders. The site name is masked like any
      name the tool shows (digits become #), and characters Windows refuses in a folder name are replaced.
    .PARAMETER OutputFolder
      The base folder.
    .PARAMETER SiteUrl
      The site address; its last part (teams/<name>) names the site folder.
    .OUTPUTS
      [string] the path.
    #>
    param([string]$OutputFolder, [string]$SiteUrl)
    $siteName = [uri]::UnescapeDataString(($SiteUrl.TrimEnd('/') -split '/')[-1])
    $siteName = ((Protect-PiiText $siteName) -replace '[\\/:*?"<>|]', '-').Trim(' .')
    if (-not $siteName) { $siteName = 'site' }
    if ($siteName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $siteName = '_' + $siteName }   # names Windows keeps for devices
    $now = Get-Date
    $folder = Join-Path (Join-Path $OutputFolder $siteName) ($now.ToString('yyyy-MM-dd'))
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
    $base = 'Ion Results ' + $now.ToString('HH-mm-ss')
    $path = Join-Path $folder ($base + '.xlsx')
    $copy = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $folder ('{0} ({1}).xlsx' -f $base, $copy); $copy++ }
    return $path
}

function Get-HitCells {
    # One row of a hit tab: the hit's cells in the order of $HitColumns.
    param($Hit)
    @($Hit.Confidence, $Hit.Channel, $Hit.Item, $Hit.MaskedValue, $Hit.Context, $Hit.Modified, $Hit.Link, $Hit.Source)
}

function Get-WorkbookSheets {
    <#
    .SYNOPSIS
      The tabs of the workbook: Summary first, then one tab per PII type with High or Medium hits (worst type
      first), then Low confidence and Not scanned when they have rows (the PAM, 2026-11-20).
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the Summary tab.
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site): the name of the folder column.
    .OUTPUTS
      [object[]] New-XlsxSheet objects.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$FolderLabel)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    $columns = @('Confidence', $FolderLabel, 'Item', 'Masked value', 'Context', 'Modified', 'Link', 'Found in')
    $widths = @(11, 22, 45, 18, 60, 17, 50, 10)
    $byRank = @{ Expression = { $script:Rank[$_.Confidence] }; Descending = $true }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-XlsxSheet 'Summary' @($script:ToolTitle, '', '', '', '') (Get-SummaryRows $Rows $ItemsScanned $Header) @(28, 70, 9, 9, 9)))
    foreach ($type in @(Get-TypeCounts $hits | Where-Object { $_.High + $_.Medium -gt 0 } | ForEach-Object { $_.Type })) {
        $typeHits = @($hits | Where-Object { $_.Detector -eq $type -and $_.Confidence -ne 'Low' } | Sort-Object $byRank, Channel, Item)
        $sheets.Add((New-XlsxSheet $type $columns @($typeHits | ForEach-Object { , (Get-HitCells $_) }) $widths -List))
    }
    $low = @($hits | Where-Object { $_.Confidence -eq 'Low' } | Sort-Object { Get-SeverityIndex $_.Detector }, Channel, Item)
    if ($low.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Low confidence' (@('Type') + $columns) @($low | ForEach-Object { , (@($_.Detector) + (Get-HitCells $_)) }) (@(16) + $widths) -List))
    }
    $notes = @($Rows | Where-Object { -not $script:Rank.ContainsKey($_.Confidence) })
    if ($notes.Count -gt 0) {
        $sheets.Add((New-XlsxSheet 'Not scanned' @('Status', $FolderLabel, 'Item', 'Reason', 'Link') `
            @($notes | ForEach-Object { , @($_.Detector, $_.Channel, $_.Item, $_.Context, $_.Link) }) @(14, 22, 45, 70, 50) -List))
    }
    return $sheets.ToArray()
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
      The rows of the Summary tab: the scan's details and totals, the hits per type (worst first), then how to read
      the tabs, the breach step and the handling note.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top (site, folders, time).
    .OUTPUTS
      [object[]] rows, each an array of cells.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header)
    $hits = @($Rows | Where-Object { $script:Rank.ContainsKey($_.Confidence) })
    foreach ($pair in @($Header.GetEnumerator()) + @((Get-ScanMetrics $Rows $ItemsScanned).GetEnumerator())) { , @($pair.Key, $pair.Value) }
    , @('')
    , @('Possible PII by type (worst first)', 'High', 'Medium', 'Low', 'Total')
    foreach ($type in @(Get-TypeCounts $hits)) { , @($type.Type, $type.High, $type.Medium, $type.Low, $type.Total) }
    if ($hits.Count -eq 0) { , @('None found in the files that were read.') }
    , @('')
    , @('Tabs', 'One tab per type with High or Medium hits, worst type first; then Low confidence; then Not scanned.')
    , @('Confidence', 'A pattern guess, not proof: open each item and check it before you act.')
    , @('If you find PII', $script:BreachAdvice)
    , @('Handling', 'Handle this file as CUI: it shows where PII is. Every value in it is masked.')
}

function Publish-ScanReport {
    <#
    .SYNOPSIS
      Shows the summary on screen and saves the workbook (Summary tab first). Returns its path.
    .PARAMETER Rows
      All report rows.
    .PARAMETER ItemsScanned
      How many files were read.
    .PARAMETER Header
      Label and value pairs for the top of the summary.
    .PARAMETER OutputFolder
      The base folder; the workbook goes in <site>\<day> under it.
    .PARAMETER SiteUrl
      The site address (names the site folder).
    .PARAMETER FolderLabel
      'Channel' (a team) or 'Folder' (a plain site).
    .OUTPUTS
      [string] the workbook path.
    #>
    param([object[]]$Rows, [int]$ItemsScanned, [System.Collections.IDictionary]$Header, [string]$OutputFolder, [string]$SiteUrl,
          [string]$FolderLabel = 'Folder')
    Get-ScanSummaryText -Rows $Rows -ItemsScanned $ItemsScanned -Header $Header -FolderLabel $FolderLabel | ForEach-Object { Write-Host $_ }
    $path = Get-ResultPath $OutputFolder $SiteUrl
    Export-Xlsx -Path $path -Sheets (Get-WorkbookSheets $Rows $ItemsScanned $Header $FolderLabel)
    Write-Host ''
    Write-Host "Saved (values masked): $path" -ForegroundColor Cyan
    Write-Host 'Handle this file as CUI: it shows where PII is.' -ForegroundColor Yellow
    return $path
}

# ==== Module ScanMeter.psm1 ====
<#
ScanMeter.psm1 - the progress bars and the per-folder lines of a scan: counts only, so the screen shows no PII.

  Purpose:     Lets the person follow a scan: files done of the total, read, not scanned, possible PII,
               time used and time left (the PAM, 2026-11-20). No file name and no found value is shown; channel
               names are masked.
  Inputs:      the report row list and the counts the scan functions of ChannelScan.psm1 give.
  Outputs:     a meter (hashtable), status text, percent, a per-file progress bar and one coloured line per folder.
  Depends on:  PiiMasking.psm1 (Protect-PiiText), imported first by Find-ChannelPII.ps1.
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     none.
  Writes:      none (screen only).
  Credentials: none.
#>


# From this size on, the per-file bar says the step can take a minute (5 MB: a slow download on a gov network
# takes tens of seconds; the author's choice, 2026-11-20).
$script:LargeFileBytes = 5MB

function New-ScanMeter {
    <#
    .SYNOPSIS
      Starts a meter for one scan step: items done, read, not scanned, possible PII, and the time.
    .PARAMETER Total
      How many items the step has (0 when not known up front).
    .OUTPUTS
      [hashtable] the meter (Update-ScanMeter and Get-ScanMeterText use it).
    #>
    param([int]$Total)
    # Clock: anything with an Elapsed [TimeSpan] (a running stopwatch here; a fixed time in the tests).
    @{ Total = $Total; Done = 0; Read = 0; Skipped = 0; Hits = 0; HitItems = 0; Clock = [Diagnostics.Stopwatch]::StartNew() }
}

function Update-ScanMeter {
    <#
    .SYNOPSIS
      Counts one item into the meter from the report rows it added.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Rows
      The report row list.
    .PARAMETER FirstRow
      The row count before the item was scanned.
    .PARAMETER ReadCount
      The items whose text was read (the scan function's return value).
    #>
    param([hashtable]$Meter, $Rows, [int]$FirstRow, [int]$ReadCount)
    $new = @(for ($r = $FirstRow; $r -lt $Rows.Count; $r++) { $Rows[$r] })
    $hits = @($new | Where-Object { $_.Confidence })   # a row with a confidence is a possible PII hit
    $Meter.Done++
    $Meter.Read += $ReadCount
    $Meter.Skipped += @($new | Where-Object { $_.Detector -eq 'NOT SCANNED' -or $_.Detector -eq 'BLOCKED' }).Count
    $Meter.Hits += $hits.Count
    $Meter.HitItems += @($hits | Select-Object -ExpandProperty Item -Unique).Count   # files with a hit
}

function Format-Duration {
    # A time as minutes:seconds; the minutes go past 59 (65 minutes = 65:00), so a long scan never wraps to 05:00.
    param([TimeSpan]$Time)
    return '{0:00}:{1:00}' -f [int][Math]::Floor($Time.TotalMinutes), $Time.Seconds
}

function Get-ScanMeterText {
    <#
    .SYNOPSIS
      The meter as one status line: counts, time used and, when the total is known, time left.
    .PARAMETER Meter
      Output of New-ScanMeter.
    .PARAMETER Noun
      What an item is, for example 'file'.
    .OUTPUTS
      [string] for example "12 of 40 files done | read 10 | not scanned 2 | possible PII 3 in 2 files | 00:42 used, about 01:48 left".
    #>
    param([hashtable]$Meter, [string]$Noun)
    $elapsed = $Meter.Clock.Elapsed
    $text = [string]$Meter.Done
    if ($Meter.Total -gt 0) { $text += ' of ' + $Meter.Total }
    $text += ' ' + $Noun + 's done | read {0} | not scanned {1} | possible PII {2} in {3} {4}s | {5} used' -f
        $Meter.Read, $Meter.Skipped, $Meter.Hits, $Meter.HitItems, $Noun, (Format-Duration $elapsed)
    if ($Meter.Total -gt 0 -and $Meter.Done -gt 0 -and $Meter.Done -lt $Meter.Total) {
        $left = [TimeSpan]::FromTicks([long]($elapsed.Ticks / $Meter.Done * ($Meter.Total - $Meter.Done)))
        $text += ', about ' + (Format-Duration $left) + ' left'
    }
    return $text
}

function Get-MeterPercent {
    # The meter's percent done (0-100); 0 when the total is not known.
    param([hashtable]$Meter)
    if ($Meter.Total -le 0) { return 0 }
    return [int][Math]::Min(100, 100 * $Meter.Done / $Meter.Total)
}

function Write-FileStep {
    <#
    .SYNOPSIS
      A second progress bar for the file in work: its size and its step, so a large file does not look hung.
      PnP gives no byte count while it downloads, so the bar moves by step (download, read, check), not by byte.
    .PARAMETER Size
      The file size in bytes.
    .PARAMETER Step
      What is being done now.
    .PARAMETER Percent
      How far through the file's steps (0-100).
    #>
    param([long]$Size, [string]$Step, [int]$Percent)
    $status = '{0} - {1:N1} MB' -f $Step, ($Size / 1MB)
    if ($Size -ge $script:LargeFileBytes) { $status += ' (a large file: this step can take a minute)' }
    Write-Progress -Id 2 -ParentId 1 -Activity 'This file' -Status $status -PercentComplete $Percent
}

function Write-ChannelLine {
    <#
    .SYNOPSIS
      Prints one line when a folder (a channel) is done: its counts, yellow when it has possible PII.
    .PARAMETER Number
      Which folder this is (1-based) and the folder count, as "2/5".
    .PARAMETER Channel
      The channel name (masked here).
    .PARAMETER Before
      A copy of the meter before the channel (Done, Read, Skipped, Hits).
    .PARAMETER Meter
      The meter after the channel.
    .PARAMETER Noun
      What an item is, for example 'file'.
    #>
    param([string]$Number, [string]$Channel, [hashtable]$Before, [hashtable]$Meter, [string]$Noun)
    $hits = $Meter.Hits - $Before.Hits
    $line = '  [{0}] {1}: {2} {3}s - read {4}, not scanned {5}, possible PII {6}' -f $Number, (Protect-PiiText $Channel),
        ($Meter.Done - $Before.Done), $Noun, ($Meter.Read - $Before.Read), ($Meter.Skipped - $Before.Skipped), $hits
    $colour = 'Green'
    if ($hits -gt 0) { $colour = 'Yellow' }
    Write-Host $line -ForegroundColor $colour
}

# ==== Module ChannelScan.psm1 ====
<#
ChannelScan.psm1 - reads one team's or site's files and turns them into report rows.

  Purpose:     The SharePoint and Teams steps of Find-ChannelPII.ps1, one function per step: read the
               Teams link or site address, sign in to the team's site, find the library, list the channel
               folders, scan each file with a progress bar and live counts. Posts are not read (the PAM,
               2026-11-20: one sign-in only; reading posts needs an admin-approved Graph permission).
  Inputs:      the Teams link or site address the person gives.
  Outputs:     report rows (Report.psm1 New-ReportRow) added to the list the caller passes in.
  Depends on:  PnP.PowerShell 1.12.0; PiiDetectors.psm1, PiiMasking.psm1, TextExtract.psm1, Report.psm1,
               ScanMeter.psm1 (all imported first by Find-ChannelPII.ps1).
  How to run:  imported by Find-ChannelPII.ps1; not run on its own.
  Run target:  Windows PowerShell 5.1 (also PowerShell 7).
  Network:     only the one SharePoint site the person gives. Nothing lists other sites or teams (no network
               reconnaissance).
  Writes:      none in SharePoint or Teams: read calls only. Rows go to an in-memory list.
  Credentials: the person's own CAC sign-in (PnP web login). Nothing is stored.
#>


# =====================================================================================
# Settings
# =====================================================================================

# A team site's file library: Teams names it "Documents" (URL "Shared Documents"); some sites renamed it.
$script:LibraryNames = @('Shared Documents', 'Documents')

# Teams makes one top-level folder per standard channel. "Forms" is SharePoint's own folder, not a channel.
$script:NotChannelFolders = @('Forms')

# How the report and the screen name the files at the top of the library (in no folder).
$script:RootGroupName = '(top of the library)'

# Items per request when listing the library. SharePoint refuses views over 5,000 items
# (list view threshold); 500 keeps each request small.
$script:ListPageSize = 500

# Accepted site address: https://<host>/sites/<name> or /teams/<name>, also as a file's "Copy link"
# (https://<host>/:w:/r/teams/<name>/...). Anything after the site name is cut off, so the scan can only
# touch that one site. Other sharing links (/:w:/s/, /:w:/t/) are refused: they do not name the site.
$script:SiteUrlPattern = '^(?<host>https://[A-Za-z0-9.-]+)/(?::[a-z]:/r/)?(?<site>(?:sites|teams)/[^/?#\s]+)'

# Accepted Teams link: a team's or a channel's "Copy link" (https://<teams-host>/l/team/... or /l/channel/...).
# Only its groupId (the team's Microsoft 365 group) is used; the host is not, so any Teams cloud's link works.
# The path may hold spaces (a channel name) but not a second address ('://' or '/l/'), so text pasted after the link is not read.
$script:TeamsLinkPattern = '^https://[A-Za-z0-9.-]+/l/(?<kind>team|channel)/(?:(?!://|/l/)[^?#\r\n])+\?(?:[^#\s]*&)?groupId=' +
                           '(?<id>[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})(?![0-9A-Fa-f-])'

# =====================================================================================
# Site, library and channels
# =====================================================================================

function Add-UrlScheme {
    <#
    .SYNOPSIS
      Adds https:// to an address typed without a scheme (usaf.dps.mil/teams/<team-site>); other text stays as it is.
    .PARAMETER Url
      The address the person gave.
    .OUTPUTS
      [string] the trimmed address, with https:// when it had no scheme.
    #>
    param([string]$Url)
    $Url = $Url.Trim()
    if ($Url -and $Url -notmatch '^[A-Za-z][A-Za-z0-9+.-]*://') { return 'https://' + $Url }
    return $Url
}

function Test-TeamSiteUrl {
    <#
    .SYNOPSIS
      Checks that the address is one SharePoint site and returns it cut to the site name.
    .PARAMETER Url
      The address the person gave (it may go on to a library or a file, or be a file's "Copy link"; https:// may be left out).
    .OUTPUTS
      [string] https://<host>/sites/<name> or https://<host>/teams/<name>
    #>
    param([string]$Url)
    $match = [regex]::Match((Add-UrlScheme $Url), $script:SiteUrlPattern, 'IgnoreCase')
    if (-not $match.Success) {
        throw "Not a team site address: '$Url'. Expected https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>."
    }
    return $match.Groups['host'].Value + '/' + $match.Groups['site'].Value
}

function Get-ScanTarget {
    <#
    .SYNOPSIS
      Reads what the person gave: a Teams link (team or channel) or a team site address.
      Throws when it is neither.
    .PARAMETER Address
      The Teams link or the site address.
    .OUTPUTS
      [pscustomobject] GroupId ([guid] from a Teams link, else $null), SiteUrl (the site, else '') and
      IsChannel (true for a channel's link).
    #>
    param([string]$Address)
    $link = [regex]::Match((Add-UrlScheme $Address), $script:TeamsLinkPattern, 'IgnoreCase')
    if ($link.Success) {
        return [pscustomobject]@{ GroupId = [guid]$link.Groups['id'].Value; SiteUrl = ''; IsChannel = ($link.Groups['kind'].Value -ieq 'channel') }
    }
    if ($Address -match '/l/(team|channel)/') {
        throw 'This Teams link has no groupId. In Teams: the ... next to the team or channel name > Copy link.'
    }
    return [pscustomobject]@{ GroupId = $null; SiteUrl = (Test-TeamSiteUrl $Address); IsChannel = $false }
}

function Assert-SiteGroup {
    <#
    .SYNOPSIS
      Checks that the connected site belongs to the team in the Teams link; throws when not (Connect-TeamSite
      then signs out), so nothing of another site is read.
    .PARAMETER GroupId
      The group ID from the Teams link.
    #>
    param([guid]$GroupId)
    $actual = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    if ($actual -ne $GroupId) {
        throw 'This site does not belong to the team in the Teams link. Nothing was scanned.'
    }
}

function Connect-TeamSite {
    <#
    .SYNOPSIS
      Signs in (own CAC, PnP web login) to the team's site and returns its address. For a Teams link it asks
      for the site address (a lookup at the tenant root hung on DoD365: the PAM, 2026-11-20), then checks that
      the site belongs to the team in the link. On any error or cancel it signs out.
    .PARAMETER Target
      Output of Get-ScanTarget.
    .OUTPUTS
      [string] the site the scan may touch.
    #>
    param($Target)
    $site = $Target.SiteUrl
    $ok = $false   # set only at the end: a cancel (Ctrl+C, Cancel on a prompt) skips 'catch' but not 'finally'
    try {
        if ($null -ne $Target.GroupId) {
            if ($Target.IsChannel) {
                Write-Host 'Channel link: the scan covers the standard channels of its team. A private or shared channel has its own site: give that site address instead.' -ForegroundColor Yellow
            }
            Write-Host "Teams link: the script needs this team's site address. In Teams: the team's Shared tab > ... > Open in SharePoint, then copy the address." -ForegroundColor Yellow
            $site = Test-TeamSiteUrl (Read-Host 'Team site address (https://usaf.dps.mil/teams/<team-site>)')
        }
        Write-Host "Signing in to $site (a sign-in window opens; if you do not see it, press Alt+Tab) ..."
        # PnP prints a warning that suggests -Interactive (and its update notice): not useful to the person, so hidden (the PAM, 2026-11-20).
        Connect-PnPOnline -Url $site -UseWebLogin -WarningAction SilentlyContinue
        if ($null -ne $Target.GroupId) { Assert-SiteGroup $Target.GroupId }
        $ok = $true
        return $site
    }
    finally {
        # No scan starts after an error or a cancel here: sign out, so no connection stays open in the ISE session.
        if (-not $ok) {
            try { Disconnect-PnPOnline } catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error
        }
    }
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
      Lists the channel folders: the top-level folders of the library, without SharePoint's own. The scan
      covers all of them (the PAM, 2026-11-20: no channel choice). A plain SharePoint site may have none.
    .PARAMETER Library
      Output of Get-TeamDocumentLibrary.
    .OUTPUTS
      [string[]] folder names, sorted.
    #>
    param($Library)
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $Library.SiteRelativeUrl -ItemType Folder -ErrorAction Stop
    return @($folders | ForEach-Object { $_.Name } | Where-Object { $script:NotChannelFolders -notcontains $_ } | Sort-Object)
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

function Get-RootFiles {
    # The files at the top of the library, in no folder (a plain SharePoint site keeps files there; a team has none).
    param([object[]]$AllItems, [string]$LibraryUrl)
    $prefix = $LibraryUrl + '/'
    $AllItems | Where-Object {
        $ref = Get-FieldValue $_ 'FileRef'
        (Get-FieldValue $_ 'FSObjType') -eq '0' -and $ref.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            $ref.Substring($prefix.Length).IndexOf('/') -lt 0
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
    try {
        Write-FileStep $size 'Downloading' 10
        $bytes = Get-FileBytes $ref
        Write-FileStep $size 'Reading the text' 50
        $text = Get-TextFromBytes -Bytes $bytes -FileName $name
    }
    catch {
        # One file that cannot be downloaded must not stop the scan: note it and go on.
        $reason = Protect-PiiText ('Download failed: ' + $_.Exception.Message)
        $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' $reason)); return 0
    }
    if ($text.Status -ne 'OK') { $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified 'NOT SCANNED' '' '' (Protect-PiiText $text.Reason))); return 0 }
    Write-FileStep $size 'Checking for PII' 80
    foreach ($hit in @(Find-PiiInText $text.Text)) {
        $Rows.Add((New-ReportRow 'File' $Channel $item $link $modified $hit.Detector $hit.Confidence $hit.MaskedValue $hit.Context))
    }
    return 1
}

function Invoke-FileScan {
    <#
    .SYNOPSIS
      Scans every file in the given channel folders (all of the team's) and adds report rows.
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
    # One group per folder; files at the top of the library (a plain SharePoint site) are one more group.
    $groups = @(foreach ($channel in $Channels) {
        $prefix = $Library.ServerRelativeUrl + '/' + $channel + '/'
        [pscustomobject]@{ Name = $channel; Prefix = $prefix; Files = @(Get-ChannelFiles $allItems $prefix) }
    })
    $rootFiles = @(Get-RootFiles $allItems $Library.ServerRelativeUrl)
    if ($rootFiles.Count -gt 0) {
        $groups += [pscustomobject]@{ Name = $script:RootGroupName; Prefix = $Library.ServerRelativeUrl + '/'; Files = $rootFiles }
    }
    $total = 0
    foreach ($group in $groups) { $total += $group.Files.Count }
    Write-Host ('Found {0} files in {1} folders. Scanning ...' -f $total, $groups.Count) -ForegroundColor Cyan
    $meter = New-ScanMeter $total
    for ($c = 0; $c -lt $groups.Count; $c++) {
        $group = $groups[$c]
        $shownChannel = Protect-PiiText $group.Name   # the progress bar is on screen too
        $before = $meter.Clone()
        foreach ($file in $group.Files) {
            Write-Progress -Id 1 -Activity ("Scanning files - folder {0} of {1}: '{2}'" -f ($c + 1), $groups.Count, $shownChannel) `
                -Status (Get-ScanMeterText $meter 'file') -PercentComplete (Get-MeterPercent $meter)
            $firstRow = $Rows.Count
            $read = Invoke-OneFileScan $file $group.Name $group.Prefix $origin $MaxFileMB $Rows
            Write-Progress -Id 2 -Activity 'This file' -Completed
            Update-ScanMeter $meter $Rows $firstRow $read
        }
        Write-ChannelLine ('{0}/{1}' -f ($c + 1), $groups.Count) $group.Name $before $meter 'file'
    }
    Write-Progress -Id 1 -Activity 'Scanning files' -Completed
    Write-Host ('Files done: ' + (Get-ScanMeterText $meter 'file')) -ForegroundColor Cyan
    return $meter.Read
}

# =====================================================================================
# Team or plain site
# =====================================================================================

function Test-TeamSite {
    <#
    .SYNOPSIS
      True when the connected site belongs to a team (a Microsoft 365 group): its folders are then channels.
      The report names the folder column "Channel" for a team and "Folder" for a plain site (the PAM, 2026-11-20).
    .OUTPUTS
      [bool]
    #>
    $groupId = (Get-PnPSite -Includes GroupId -ErrorAction Stop).GroupId
    return [bool]($groupId -and $groupId -ne [guid]::Empty)
}

# ==== Body ==== (scripts/build_paste.py copies everything from this line into the paste bundle; the lines above
# become the bundle's Settings block and its inline module code)

Write-Host 'Ion - DoD365 PII Reporter (read-only). Values in the results are masked.' -ForegroundColor Cyan
$clock = [Diagnostics.Stopwatch]::StartNew()

# ---- Step 1: the Teams link or site address (only this one site is touched) ----
if (-not $SiteUrl) { $SiteUrl = Read-Host 'Teams link, or SharePoint site address (https://usaf.dps.mil/teams/<team-site> or /sites/<team-site>)' }
$target = Get-ScanTarget $SiteUrl

# ---- Step 2: sign in with the person's own CAC (a Teams link first asks for the team's site) ----
$SiteUrl = Connect-TeamSite $target

$rows = New-Object System.Collections.Generic.List[object]
$scanned = 0
$selected = @()
$folderLabel = 'Folder'   # 'Channel' when the site belongs to a team (the PAM, 2026-11-20)
$finished = $false   # set only after the last step; a cancelled run (Ctrl+C) skips 'catch' but not 'finally'
try {
    # ---- Step 3: the folders to scan: all of them (in a team, one per channel) ----
    # Only the column name depends on it: if the site's group cannot be read, the scan goes on with 'Folder'.
    try { if (Test-TeamSite) { $folderLabel = 'Channel' } }
    catch { Write-Verbose 'Team check not possible: the folder column is named Folder.' }
    $library = Get-TeamDocumentLibrary
    $selected = @(Get-ChannelFolderNames $library)
    $shownFolders = ($selected | ForEach-Object { Protect-PiiText $_ }) -join ', '
    Write-Host ($folderLabel + 's to scan: ' + $shownFolders + '; also any files at the top of the library')

    # ---- Step 4: the files ----
    $scanned += Invoke-FileScan -Library $library -Channels $selected -SiteUrl $SiteUrl -MaxFileMB $MaxFileMB -Rows $rows
    $finished = $true
}
catch {
    # Keep what was found so far: the results say the scan stopped early, and why (PII in the message masked).
    $reason = Protect-PiiText $_.Exception.Message
    Write-Host ('The scan stopped early: ' + $reason) -ForegroundColor Red
    $rows.Add((New-ReportRow 'Scan' '(all)' '(scan)' '' '' 'INCOMPLETE' '' '' ('Stopped early: ' + $reason)))
}
finally {
    # ---- Step 5: the results (also after an error or a cancel), then sign out ----
    Add-IncompleteRow -Rows $rows -Finished $finished
    $header = [ordered]@{}
    $header['Site'] = $SiteUrl
    $header[$folderLabel + 's'] = ($selected | ForEach-Object { Protect-PiiText $_ }) -join ', '
    $header['Run'] = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $header['Time used'] = '{0:00}:{1:00}' -f [int][Math]::Floor($clock.Elapsed.TotalMinutes), $clock.Elapsed.Seconds
    try {
        $null = Publish-ScanReport -Rows $rows.ToArray() -ItemsScanned $scanned -Header $header -OutputFolder $OutputFolder `
            -SiteUrl $SiteUrl -FolderLabel $folderLabel
    }
    finally {
        # Sign out even when the results could not be saved (for example, a folder that cannot be written).
        try { Disconnect-PnPOnline }
        catch { Write-Verbose 'Already signed out.' }  # nothing to sign out of: not an error for the person
    }
}

# ==== Clean-up ====
} finally {
    Set-StrictMode -Off
    $ErrorActionPreference = $previousErrorAction
}
