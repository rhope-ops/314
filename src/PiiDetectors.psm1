<#
PiiDetectors.psm1 - finds likely PII in plain text and masks what it finds.

  Purpose:     Pattern detectors (SSN, DoD ID, date of birth, payment card, phone, personal e-mail,
               passport). Each hit gets a confidence label and a MASKED value; the full value never
               leaves this module.
  Network:     none.
  Writes:      none (pure functions on strings).
  Credentials: none.
  Runs on:     Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version 2.0

# Keyword windows: how many characters before a match we look at for a label such as "SSN".
$script:KeywordWindow = 40

# Detector order = priority. When two hits overlap, the earlier detector in this list wins.
$script:DetectorOrder = @('SSN', 'DoD ID', 'Payment card', 'Date of birth', 'Passport', 'Phone', 'Personal e-mail')

$script:ConfidenceRank = @{ 'High' = 3; 'Medium' = 2; 'Low' = 1 }

function Get-ConfidenceRank {
    # Returns 3/2/1 for High/Medium/Low so callers can sort and pick the highest label.
    param([string]$Confidence)
    if ($script:ConfidenceRank.ContainsKey($Confidence)) { return $script:ConfidenceRank[$Confidence] }
    return 0
}

function Test-KeywordBefore {
    <#
    .SYNOPSIS  True when a keyword pattern appears in the text just before (or around) a match.
    .PARAMETER Text      The full text.
    .PARAMETER Index     Start index of the match.
    .PARAMETER Pattern   Case-insensitive regex of keywords.
    .PARAMETER Window    How many characters before the match to search.
    #>
    param([string]$Text, [int]$Index, [string]$Pattern, [int]$Window = $script:KeywordWindow)
    $start = [Math]::Max(0, $Index - $Window)
    $before = $Text.Substring($start, $Index - $start)
    return [regex]::IsMatch($before, $Pattern, 'IgnoreCase')
}

function Test-SsnValid {
    <#
    .SYNOPSIS  Applies the SSA issuing rules: area not 000, 666 or 900-999; group not 00; serial not 0000.
    #>
    param([string]$Area, [string]$Group, [string]$Serial)
    $a = [int]$Area
    if ($a -eq 0 -or $a -eq 666 -or $a -ge 900) { return $false }
    if ([int]$Group -eq 0) { return $false }
    if ([int]$Serial -eq 0) { return $false }
    return $true
}

function Test-Luhn {
    <#
    .SYNOPSIS  Luhn (mod 10) checksum used by payment card numbers.
    .PARAMETER Digits  A string of digits only.
    #>
    param([string]$Digits)
    $sum = 0
    $double = $false
    for ($i = $Digits.Length - 1; $i -ge 0; $i--) {
        $d = [int][string]$Digits[$i]
        if ($double) {
            $d = $d * 2
            if ($d -gt 9) { $d = $d - 9 }
        }
        $sum += $d
        $double = -not $double
    }
    return ($sum % 10 -eq 0)
}

function Get-MaskedValue {
    <#
    .SYNOPSIS  Masks a found value so the report never holds the full PII.
    .DESCRIPTION
      SSN       -> ***-**-1234      DoD ID   -> ******7890
      Card      -> ************1234 Phone    -> ***-***-1234
      E-mail    -> ***@domain       Passport -> *****123
      DOB       -> [date masked]
    #>
    param([string]$Detector, [string]$Value)
    $digits = ($Value -replace '\D', '')
    switch ($Detector) {
        'SSN' { return '***-**-' + $digits.Substring($digits.Length - 4) }
        'DoD ID' { return '******' + $digits.Substring($digits.Length - 4) }
        'Payment card' { return ('*' * ($digits.Length - 4)) + $digits.Substring($digits.Length - 4) }
        'Phone' { return '***-***-' + $digits.Substring($digits.Length - 4) }
        'Personal e-mail' { return '***@' + ($Value -split '@')[-1] }
        'Passport' {
            $keep = [Math]::Min(3, $Value.Length)
            return ('*' * ($Value.Length - $keep)) + $Value.Substring($Value.Length - $keep)
        }
        'Date of birth' { return '[date masked]' }
        default { return '[masked]' }
    }
}

function Get-MaskedContext {
    <#
    .SYNOPSIS  Returns a short snippet around a hit with every digit replaced by '#' and every
               e-mail local part replaced by '***', so the snippet cannot leak the value.
    #>
    param([string]$Text, [int]$Index, [int]$Length, [int]$Radius = 30)
    $start = [Math]::Max(0, $Index - $Radius)
    $end = [Math]::Min($Text.Length, $Index + $Length + $Radius)
    $snip = $Text.Substring($start, $end - $start)
    $snip = $snip -replace '[A-Za-z0-9._%+-]+@', '***@'
    $snip = $snip -replace '\d', '#'
    $snip = $snip -replace '\s+', ' '
    return $snip.Trim()
}

function New-PiiHit {
    # Builds one hit object. Value is kept only inside this module for overlap checks; it is
    # removed before the hit is returned (see Find-PiiInText).
    param([string]$Detector, [string]$Confidence, [System.Text.RegularExpressions.Match]$Match, [string]$Value)
    [pscustomobject]@{
        Detector   = $Detector
        Confidence = $Confidence
        Index      = $Match.Index
        Length     = $Match.Length
        Value      = $Value
    }
}

function Find-SsnHits {
    <#
    .SYNOPSIS  Social Security numbers: 3-2-4 digits with the same separator (dash, space or none).
    .DESCRIPTION
      High   = a keyword (SSN, SSAN, social security) is near.
      Medium = dashes or spaces, no keyword.
      Low    = 9 bare digits, no keyword (many false positives; review by hand).
    #>
    param([string]$Text)
    $kw = 'SSN|SSAN|social\s*security|soc\.?\s*sec'
    foreach ($m in [regex]::Matches($Text, '(?<![\d-])(\d{3})([- ]?)(\d{2})\2(\d{4})(?![\d-])')) {
        if (-not (Test-SsnValid $m.Groups[1].Value $m.Groups[3].Value $m.Groups[4].Value)) { continue }
        $conf = 'Low'
        if ($m.Groups[2].Value -ne '') { $conf = 'Medium' }
        if (Test-KeywordBefore $Text $m.Index $kw) { $conf = 'High' }
        New-PiiHit 'SSN' $conf $m $m.Value
    }
}

function Find-DodIdHits {
    <#
    .SYNOPSIS  DoD ID (EDIPI): 10 digits, reported only when a label (DoD ID, DODID, EDIPI, CAC) is near.
    #>
    param([string]$Text)
    $kw = 'DoD\s*ID|DODID|EDIPI|\bCAC\b|DoD\s*(?:ID\s*)?(?:#|no\.?|number)'
    foreach ($m in [regex]::Matches($Text, '(?<!\d)\d{10}(?!\d)')) {
        if (Test-KeywordBefore $Text $m.Index $kw) { New-PiiHit 'DoD ID' 'High' $m $m.Value }
    }
}

function Find-CardHits {
    <#
    .SYNOPSIS  Payment card numbers: 13-19 digits (spaces or dashes allowed) that pass the Luhn check.
    .DESCRIPTION  High with a keyword (card, Visa, Mastercard, Amex, GTC), else Medium.
    #>
    param([string]$Text)
    $kw = 'card|visa|master\s*card|amex|american\s*express|discover|\bGTC\b|\bGPC\b'
    foreach ($m in [regex]::Matches($Text, '(?<![\d-])\d(?:[ -]?\d){12,18}(?![\d-])')) {
        $digits = $m.Value -replace '\D', ''
        if ($digits.Length -lt 13 -or -not (Test-Luhn $digits)) { continue }
        $conf = 'Medium'
        if (Test-KeywordBefore $Text $m.Index $kw) { $conf = 'High' }
        New-PiiHit 'Payment card' $conf $m $m.Value
    }
}

function Find-DobHits {
    <#
    .SYNOPSIS  Dates reported as date of birth only when a label (DOB, date of birth, born) is near.
    .DESCRIPTION  Formats: 1/2/1990, 01-02-90, 1990-01-02, 2 Jan 1990, 02JAN1990, January 2, 1990.
    #>
    param([string]$Text)
    $kw = '\bDOB\b|date\s*of\s*birth|birth\s*date|birthdate|\bborn\b|\bD\.O\.B\.?'
    $mon = '(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*'
    $date = '\b(?:\d{1,2}[/-]\d{1,2}[/-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2}|\d{1,2}\s?' + $mon +
            '\s?\d{2,4}|' + $mon + '\s+\d{1,2},?\s+\d{4})\b'
    foreach ($m in [regex]::Matches($Text, $date, 'IgnoreCase')) {
        if (Test-KeywordBefore $Text $m.Index $kw 30) { New-PiiHit 'Date of birth' 'High' $m $m.Value }
    }
}

function Find-PassportHits {
    <#
    .SYNOPSIS  Passport numbers: 6-9 letters/digits (at least 5 digits) with the word "passport" near.
    #>
    param([string]$Text)
    foreach ($m in [regex]::Matches($Text, '\b[A-Z0-9]{6,9}\b')) {
        if (($m.Value -replace '\D', '').Length -lt 5) { continue }
        if (Test-KeywordBefore $Text $m.Index 'passport' 30) { New-PiiHit 'Passport' 'Medium' $m $m.Value }
    }
}

function Find-PhoneHits {
    <#
    .SYNOPSIS  US phone numbers (10 digits). Medium with a personal label (cell, mobile, home, personal), else Low.
    #>
    param([string]$Text)
    $kw = '\bcell\b|mobile|\bhome\b|personal|\bmy\s+(?:number|phone)'
    foreach ($m in [regex]::Matches($Text, '(?<![\d-])(?:\+?1[ .-]?)?(?:\(\d{3}\)\s?|\d{3}[ .-])\d{3}[ .-]\d{4}(?![\d-])')) {
        $conf = 'Low'
        if (Test-KeywordBefore $Text $m.Index $kw 30) { $conf = 'Medium' }
        New-PiiHit 'Phone' $conf $m $m.Value
    }
}

function Find-PersonalEmailHits {
    <#
    .SYNOPSIS  E-mail addresses outside government domains (last label not gov or mil), Medium.
    #>
    param([string]$Text)
    $govLabels = @('gov', 'mil')
    foreach ($m in [regex]::Matches($Text, '\b[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.([A-Za-z]{2,}))\b')) {
        if ($govLabels -contains $m.Groups[2].Value.ToLower()) { continue }
        New-PiiHit 'Personal e-mail' 'Medium' $m $m.Value
    }
}

function Find-PiiInText {
    <#
    .SYNOPSIS  Runs every detector on one text and returns masked hits.
    .PARAMETER Text  Plain text (already extracted from the file or post).
    .OUTPUTS   Objects with Detector, Confidence, MaskedValue, Context. No full values.
    .NOTES     Overlapping hits: the detector earlier in $DetectorOrder wins, so one number is
               reported once (for example a 9-digit SSN inside a 16-digit card number is dropped).
    #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return }

    $all = @()
    $all += @(Find-SsnHits $Text)
    $all += @(Find-DodIdHits $Text)
    $all += @(Find-CardHits $Text)
    $all += @(Find-DobHits $Text)
    $all += @(Find-PassportHits $Text)
    $all += @(Find-PhoneHits $Text)
    $all += @(Find-PersonalEmailHits $Text)

    # Sort by detector priority, then by confidence (highest first), then keep non-overlapping hits.
    $sorted = $all | Where-Object { $_ } | Sort-Object `
        @{ Expression = { [array]::IndexOf($script:DetectorOrder, $_.Detector) } },
        @{ Expression = { Get-ConfidenceRank $_.Confidence }; Descending = $true },
        Index
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($h in $sorted) {
        $overlap = $false
        foreach ($k in $kept) {
            if ($h.Index -lt ($k.Index + $k.Length) -and $k.Index -lt ($h.Index + $h.Length)) { $overlap = $true; break }
        }
        if (-not $overlap) { $kept.Add($h) }
    }

    foreach ($h in ($kept | Sort-Object Index)) {
        [pscustomobject]@{
            Detector    = $h.Detector
            Confidence  = $h.Confidence
            MaskedValue = Get-MaskedValue $h.Detector $h.Value
            Context     = Get-MaskedContext $Text $h.Index $h.Length
        }
    }
}

Export-ModuleMember -Function Find-PiiInText, Get-MaskedValue, Get-MaskedContext, Test-Luhn, Test-SsnValid, Get-ConfidenceRank
