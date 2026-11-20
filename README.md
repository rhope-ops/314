# Teams channel PII spot check

A read-only PowerShell script that a **team owner** runs to find files (and, if allowed, posts) in one Microsoft Teams team that may hold PII. It shows a readable summary on screen and saves a readable text report, a CSV, or both (you choose). **Every found value is masked** (for example `***-**-1234`), so the report is not a PII spill itself.

> This is a spot check, not a compliance control. The standard tool for PII in Teams is **Microsoft Purview** (DLP policies and Content Search), run by your wing's privacy or cybersecurity office.

## What it scans

| Content | Scanned? | How |
|---|---|---|
| Files in standard channels (Files tab) | Yes | Your own CAC sign-in to the team's SharePoint site |
| .docx .xlsx .pptx (and macro versions), .txt .csv .md .log .json .xml .html | Yes | Text read in memory; nothing saved to disk |
| PDF, images, legacy .doc/.xls/.ppt, files over 25 MB | No | Listed as **NOT SCANNED**: check them by hand |
| Channel posts and replies (`-IncludePosts`) | Only if your tenant allows it | Needs a Microsoft Graph permission (ChannelMessage.Read.All) that a tenant admin must approve. If it is missing, the report shows **BLOCKED** and the file scan still runs |
| Private and shared channels | Run again with that channel's own site URL | Each one has its own SharePoint site |
| 1:1 chats, group chats, meeting chats | No | Admin / Purview only |

## What it looks for

| Type | Rule | Confidence |
|---|---|---|
| SSN | 3-2-4 digits, valid by SSA issuing rules | High with "SSN"/"SSAN"/"social security" near; Medium with dashes; Low for 9 bare digits |
| DoD ID (EDIPI) | 10 digits with "DoD ID", "DODID", "EDIPI" or "CAC" near | High |
| Date of birth | A date with "DOB", "date of birth" or "born" near | High |
| Payment card | 13–19 digits passing the Luhn check | High with "card"/"GTC"/"Visa" etc. near; else Medium |
| Passport | 6–9 letters/digits with "passport" near | Medium |
| Phone | US 10-digit number | Medium with "cell"/"mobile"/"home"/"personal" near; else Low |
| Personal e-mail | Any address not ending in .gov or .mil | Medium |
| Home address | House number + street name + street type (St, Ave, Rd ...), optional city, state, ZIP | High with "home address"/"residence" near; Medium with state + ZIP; Low for a street only |
| Name | A name after a label: Name:, Full name:, Member:, Spouse:, Dependent:, NOK:, POC: | High when an SSN, DoD ID, date of birth, passport or card number is near; else Low |
| PII in a file or folder name | Any of the above in the path | As above; reported with Source "File name" |

Confidence is a pattern guess. Open each item and check it before you act. A name **without** a label is not detected: that would need a list of real names.

## Run it (gov device)

Needs: Windows PowerShell 5.1 and PnP.PowerShell 1.12.0. No other modules.

1. Download this repo as a ZIP (green **Code** button > **Download ZIP**) and extract it.
2. Open PowerShell in the extracted `src` folder and run:

```powershell
Get-ChildItem . | Unblock-File
Set-ExecutionPolicy -Scope Process Bypass
.\Find-ChannelPII.ps1
```

3. Paste the team's site URL when asked. To find it: in Teams, open the channel's **Files** tab > **Open in SharePoint**, and copy the address up to the site name (`https://<tenant>.sharepoint-mil.us/sites/<team-site>`).
4. Sign in with your CAC in the window that opens.
5. Pick a channel number, several numbers (`1,3`), or `A` for all.
6. Pick what to save: **1** Both (readable `.txt` and `.csv`, press Enter), **2** Text only, **3** CSV only.
7. Read the summary on screen. The saved file paths are printed at the end (default folder: `Documents\PII-Scan`).

What each report holds:

| Report | Use it to | Holds |
|---|---|---|
| Screen and `.txt` | Read and print | Totals; items with possible PII (highest confidence first); hits per type; each hit with its masked value and masked context; items not scanned |
| `.csv` | Sort and filter in Excel | One row per hit or note: Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link |

Options:

```powershell
.\Find-ChannelPII.ps1 -SiteUrl https://<tenant>.sharepoint-mil.us/sites/<team-site> -Channel All -IncludePosts -Report Both
```

| Parameter | Meaning |
|---|---|
| `-SiteUrl` | The team's site; skip the prompt |
| `-Channel` | A channel folder name, or `All`; skip the list |
| `-IncludePosts` | Also try posts and replies (second sign-in) |
| `-Report` | `Both`, `Text` or `Csv`; skip the question |
| `-OutputFolder` | Where to save the report files |
| `-MaxFileMB` | Size limit per file (default 25) |

## Safety

- **Read-only:** it never changes, moves or deletes anything in Teams or SharePoint.
- **One site only:** it touches only the site URL you give. It never lists other sites or teams.
- **No stored credentials:** your own CAC sign-in each run.
- **Masked report:** the screen, the `.txt` and the `.csv` never hold a full value. The context snippet is cut from text that is already masked (digits become `#`, found values and e-mail names become `*`), so a cut cannot expose part of a value. PII in a file or folder name is masked too, and its link then points to the library only.
- **Safe CSV:** a cell that Excel would read as a formula gets a leading `'`.
- **Stops cleanly:** if something fails part way, the report is still saved, says INCOMPLETE and why, and the script signs out.
- **Handle the reports as CUI anyway:** file names and masked fragments still point to where PII is.
- If it finds PII, follow your wing's PII breach process (DAFI 33-332) through your Privacy Manager.

## Files

Each file starts with a header: purpose, inputs, outputs, how to run, run target, and the three safety lines `Network:`, `Writes:`, `Credentials:`. Read those before you run anything.

| File | What it does |
|---|---|
| `src/Find-ChannelPII.ps1` | Entry point: the six steps (address, sign-in, channels, files, posts, report) |
| `src/ChannelScan.psm1` | SharePoint and Teams steps: check the address, list channels, read files, try posts |
| `src/TextExtract.psm1` | Text from Office files, text files and HTML, in memory |
| `src/PiiDetectors.psm1` | One detector per PII type; finds positions and raw values |
| `src/PiiMasking.psm1` | Masks every value and snippet; the only way results leave the detectors |
| `src/Report.psm1` | Screen summary, the report choice, the `.txt` and `.csv` files |
| `tests/Run-Tests.ps1` | Runs all tests (invented data, fake PnP cmdlets, no network) |
| `tests/TestKit.ps1`, `tests/FakePnP.ps1` | The test assertion and samples; the fake PnP cmdlets |
| `tests/*.Tests.ps1` | One test file per module, plus a start-to-end run |
| `tests/Lint.ps1` | Windows PowerShell 5.1 syntax check (PSScriptAnalyzer) |

## Standards and policy it follows

| Source | What it means here |
|---|---|
| Code standard of rhope-ops/90-161.Z (kit rules C1–C11, with `code-c1-safety: on`, `code-c2: full`, `code-max-line-chars: 200`) and the RADE extras X1–X7 | Headers with safety lines and run target; full help for each function; functions ≤ 60 lines, files ≤ 600, lines ≤ 200; named values with their source; every discarded error says why; `Set-StrictMode -Version 2`. Measured with 90-161.Z's `scripts/code_standard.py`: 0 violations (2026-11-20) |
| Windows PowerShell 5.1, PnP.PowerShell 1.12.0, no other modules | Lint checks 5.1 syntax; only .NET built-ins are used |
| No network reconnaissance | Only the one site address given; nothing lists other sites or teams |
| Government data stays in DoD365 | File content is read in memory only; reports stay on your computer; the tests use invented data only |
| Credentials | Your own CAC sign-in each run; nothing stored |
| No scripts in results | Script files are not read; only their masked names can appear in a report |
| Decisions in comments | Each names who decided and the date (for example "the PFM, 2026-11-20") |

## Develop

```bash
pwsh -NoProfile -File tests/Run-Tests.ps1
```

```bash
pwsh -NoProfile -File tests/Lint.ps1
```

Pass = `TESTS ... failed=0` and `LINT ... findings=0`. PSScriptAnalyzerSettings.psd1 says which style rules are off and why. GitHub Actions runs both on every push and pull request. Test data is invented; PII-shaped samples are built from parts at run time so no complete PII-shaped string is stored in this public repo.
