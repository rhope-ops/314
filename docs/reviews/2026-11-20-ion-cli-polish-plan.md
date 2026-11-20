Approved by the PAM 2026-11-20 (plan dialog). Steps 1-8 done 2026-11-20 (PR #13).

# Ion CLI polish (plan, 2026-11-20)

## Context
The PAM ran Ion on DoD365 and reported: live lines jump in width; too much gray; the 2nd progress bar is indented;
the site/org folder name is masked (###); tables and links stop at the folder, so the exact file cannot be opened;
many spreadsheets fail with "Exception calling .ctor ... End of central directory not found".
Done = all items below built, `scripts/check.sh` ends `CHECK PASS`, fresh review has no MAJOR, PR merged, paste bundle rebuilt.

DEC (the PAM, 2026-11-20, dialog; scope: the tool): space-padded counts; mask only detected PII in names + direct
file links; Cyan + Green added to the palette; all 4 extras. Replaces in part the DECISIONS.md rows on FULCRUM screen
colours and "every digit in names masked". Project rule "Masked output" (found PII values) stays fully kept.

## Changes
1. **Fixed-width live line** (`src/ScanMeter.psm1` Write-MeterLine, Get-ScanMeterText, Write-ChannelLine):
   counts right-aligned in 5 characters (`{0,5}`), `done/total` as `{0,5}/{1,5}`, percent `{0,3}%`; the
   variable part (PII per type, time) goes last. Same format in the Write-Progress status text.
2. **2nd bar alignment**: remove `-ParentId 1` from Write-FileStep (ISE indents child bars). Update `tests/FakePnP.ps1` expectation.
3. **Colours** (`src/Report.psm1` $ScreenColors + Write-Ui ValidateSet): add `Label`=Cyan, `Good`=Green,
   `Rule`=DarkGray; Text/Title White; Accent Yellow; Alert Red. Zero counts White (not Gray); 0 PII = Green;
   "Saved", "done" lines Green; banner lines (`==== ... ====`) for start, each phase and the summary; a DarkGray
   separator after each folder line. Gray stays only for the hint lines.
4. **Names**: new `Protect-PiiName` in `src/PiiMasking.psm1` = mask detected spans + e-mail local parts, keep other
   digits. Use it for site folder (Report.psm1 Get-ResultPath), channel/folder names, file paths (ChannelScan.psm1
   Invoke-OneFileScan, progress bar). Protect-PiiText (strict) stays for error messages.
5. **Exact file link** (`src/ChannelScan.psm1` Get-SafeLink): full path link when the path has no detected PII; else
   for Office files `<site>/_layouts/15/Doc.aspx?sourcedoc={UniqueId}` (add `UniqueId` to the list fields); else the folder.
   The screen table "Item" column keeps the full masked path incl. file name (verify it is not cut).
6. **Zip error** (`src/TextExtract.psm1` Get-TextFromBytes): check the first bytes before ZipArchive:
   0 bytes = "Empty file"; `D0 CF 11 E0` = "Encrypted, password-protected or sensitivity-labelled (or an old format
   with a new name): check by hand"; starts with `<` = "Download gave a web page, not the file"; else "Not a valid
   Office file". Likely cause: CUI labels with encryption store the file as OLE, not zip (Confidence: medium; the new
   reason text proves it on the next run). Content of encrypted files cannot be read without the MIP SDK: out of scope.
7. **Extras**
   - Clickable links: `src/XlsxWriter.psm1` writes `<hyperlinks>` with sheet relationships (no formula; safety rule kept).
   - Not-scanned reasons counted: Report.psm1 groups NOT SCANNED rows by reason category; on Summary tab and screen.
   - Open at end: Find-ChannelPII.ps1 asks "Open the workbook now? [Y/n]" then `Invoke-Item` (local file only).
   - File on bar: Write-Progress -Id 1 CurrentOperation = masked file path (Protect-PiiName).
8. Tests: update `tests/Detectors.Tests.ps1` (FY27 stays FY27 under Protect-PiiName; strict cases stay for
   Protect-PiiText), add tests for the byte sniff, link choice, hyperlink XML, reason counts, padded line.
   Rebuild `paste/33332Z-pii-spot-check.ps1` with `scripts/build_paste.py`. Record the DEC row in docs/DECISIONS.md; update README if it names colours/masking; log row in the private log.

## Verification
- `scripts/check.sh` → `CHECK PASS` (Run-Tests.ps1, Lint findings=0, bundle test).
- pwsh FakePnP end-to-end run: capture screen text, check widths of 3 live lines are equal; one `.xlsx` opened
  with a zip read check for `<hyperlinks>`; invented encrypted (OLE header) and 0-byte samples give the new reasons.
- Not verifiable here: ISE rendering and colours on the gov device; the PAM tests one DoD365 run.

## Delivery
Branch `claude/exciting-shannon-mlq8mc`, one PR, lane B (UI approved in this dialog): fresh `reviewer` subagent, merge on green.
