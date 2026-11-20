<#
Detectors.Tests.ps1 - tests of PiiDetectors.psm1 and PiiMasking.psm1: what is found, what is not, and masking.

  Purpose:     True and false cases for each PII type; overlap; masked values and snippets never
               show a full value.
  Inputs:      $Sample (TestKit.ps1).
  Outputs:     passes and failures counted by Assert-Equal.
  How to run:  dot-sourced by tests/Run-Tests.ps1.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none.
#>

# ---- SSN ----
Assert-Equal 'SSN/High/***-**-6789' (Get-HitSummary "Member SSN: $($Sample.Ssn) on file") 'SSN with label'
Assert-Equal 'SSN/Medium/***-**-6789' (Get-HitSummary "ref $($Sample.Ssn) end") 'SSN with dashes, no label'
Assert-Equal 'SSN/Low/***-**-6789' (Get-HitSummary "value $($Sample.SsnBare) end") 'SSN bare digits'
foreach ($bad in @('SsnArea000', 'SsnArea666', 'SsnArea9xx', 'SsnGroup00', 'SsnSerial0', 'SsnMixed')) {
    Assert-Equal 0 (Get-HitSummary "SSN $($Sample[$bad])").Count "SSN refused: $bad"
}

# ---- DoD ID ----
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary "DoD ID: $($Sample.DodId)") 'DoD ID with label'
Assert-Equal 'DoD ID/High/******7890' (Get-HitSummary "EDIPI $($Sample.DodId)") 'EDIPI with label'
Assert-Equal 0 (Get-HitSummary "tracking $($Sample.DodId) shipped").Count '10 digits without label not reported'

# ---- Payment card ----
Assert-Equal 'Payment card/High/************1111' (Get-HitSummary 'GTC card 4111 1111 1111 1111 exp') 'Card with label'
Assert-Equal 'Payment card/Medium/************4444' (Get-HitSummary 'number 5555-5555-5555-4444') 'Card without label'
Assert-Equal 0 (Get-HitSummary 'number 4111 1111 1111 1112').Count 'Card failing Luhn not reported'
Assert-Equal 1 (Get-HitSummary 'card 4111111111111111').Count 'Card digits not also reported as SSN (overlap)'

# ---- Date of birth ----
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'DOB: 03/14/1985') 'DOB numeric'
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'Date of birth 14 Mar 1985') 'DOB day month year'
Assert-Equal 'Date of birth/High/[date masked]' (Get-HitSummary 'born March 14, 1985') 'DOB month name'
Assert-Equal 0 (Get-HitSummary 'Suspense 03/14/2027').Count 'Date without birth label not reported'

# ---- Passport, phone, e-mail ----
Assert-Equal 'Passport/Medium/*****567' (Get-HitSummary "Passport no. $($Sample.Passport)") 'Passport with label'
Assert-Equal 'Phone/Medium/***-***-0142' (Get-HitSummary "my cell $($Sample.Phone)") 'Phone with personal label'
Assert-Equal 'Phone/Low/***-***-0142' (Get-HitSummary "office $($Sample.PhoneParen)") 'Phone without label'
Assert-Equal 'Personal e-mail/Medium/***@example.com' (Get-HitSummary "write to $($Sample.Email)") 'Personal e-mail'
Assert-Equal 0 (Get-HitSummary $Sample.GovEmail).Count '.gov e-mail not reported'
Assert-Equal 0 (Get-HitSummary $Sample.MilEmail).Count '.mil e-mail not reported'

# ---- Home address ----
Assert-Equal 'Home address/High/[address masked]' (Get-HitSummary "Home address: $($Sample.Street), $($Sample.CityZip)") 'Address with label'
Assert-Equal 'Home address/Medium/[address masked]' (Get-HitSummary "send to $($Sample.Street), $($Sample.CityZip) today") 'Address with ZIP'
Assert-Equal 'Home address/Low/[address masked]' (Get-HitSummary "meet at $($Sample.Street) today") 'Street only'

# ---- Name ----
Assert-Equal 'Name/Low/P*** Q*** S***' (Get-HitSummary 'Name: Pat Q. Sample') 'Labelled name'
Assert-Equal 'Name/High/S*** P*** Q***' ((Get-HitSummary "NOK: SAMPLE, PAT Q. SSN $($Sample.Ssn)")[0]) 'Name next to SSN is High'
Assert-Equal 'Name/Low/P*** S***' (Get-HitSummary 'Spouse: Pat Sample, Phone ext 4') 'Name stops at the next label'
Assert-Equal 0 (Get-HitSummary 'Pat Sample wrote the brief').Count 'Name without label not reported'

# ---- Masking: snippets never show digits, e-mail names or found names ----
$hits = @(Find-PiiInText "SSN $($Sample.Ssn) then contact person: $($Sample.Email) later")
Assert-Equal $false ($hits[0].Context -match '\d') 'Snippet has no digits'
Assert-Equal $false ($hits[0].Context -match 'pat|sample') 'Snippet cut through an e-mail shows no part of its name'
$quoted = @(Find-PiiInText "SSN $($Sample.Ssn) mail $($Sample.EmailQuote)")
Assert-Equal $false (($quoted | ForEach-Object { $_.Context }) -join ' ' -match 'connor|jane') "Snippet hides an e-mail name with an apostrophe"
$named = @(Find-PiiInText "Name: Pat Q. Sample SSN $($Sample.Ssn)")
Assert-Equal $false (($named | ForEach-Object { $_.Context }) -join ' ' -match 'Pat|Sample') 'Snippet hides a found name'

# ---- Protect-PiiText: short texts shown as they are ----
Assert-Equal 'Roster ***-**-6789.docx' (Protect-PiiText "Roster $($Sample.Ssn).docx") 'File name with SSN masked'
Assert-Equal 'FY27 Roster.docx' (Protect-PiiText 'FY27 Roster.docx') 'File name without PII unchanged'

# ---- Validators ----
Assert-Equal $true (Test-Luhn '4111111111111111') 'Luhn true'
Assert-Equal $false (Test-Luhn '4111111111111112') 'Luhn false'
