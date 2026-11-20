# Teams channel PII spot check

A read-only PowerShell script that a **team owner** runs to find files (and, if allowed, posts) in one Microsoft Teams team that may hold PII. It prints a summary table and saves a CSV. **Every found value is masked** (for example `***-**-1234`), so the report is not a PII spill itself.

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

Confidence is a pattern guess. Open each item and check it before you act. Names and home addresses are **not** detected (too many false matches without a name list).

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
6. Read the summary. The CSV path is printed at the end (default: `Documents\PII-Scan`).

Options:

```powershell
.\Find-ChannelPII.ps1 -SiteUrl https://<tenant>.sharepoint-mil.us/sites/<team-site> -Channel All -IncludePosts
```

| Parameter | Meaning |
|---|---|
| `-SiteUrl` | The team's site; skip the prompt |
| `-Channel` | A channel folder name, or `All`; skip the list |
| `-IncludePosts` | Also try posts and replies (second sign-in) |
| `-OutputFolder` | Where to save the CSV |
| `-MaxFileMB` | Size limit per file (default 25) |

## Safety

- **Read-only:** it never changes, moves or deletes anything in Teams or SharePoint.
- **One site only:** it touches only the site URL you give. It never lists other sites or teams.
- **No stored credentials:** your own CAC sign-in each run.
- **Masked report:** the CSV and console never hold a full value. Context snippets replace every digit with `#`.
- **Handle the CSV as CUI anyway:** file names and masked fragments still point to where PII is.
- If it finds PII, follow your wing's PII breach process (DAFI 33-332) through your Privacy Manager.

## Files

| File | What it does |
|---|---|
| `src/Find-ChannelPII.ps1` | Entry point: prompts, sign-in, scan, report |
| `src/ChannelScan.psm1` | SharePoint and Teams steps (list channels, read files, read posts) |
| `src/TextExtract.psm1` | Text from Office files, text files and HTML, in memory |
| `src/PiiDetectors.psm1` | The detectors and the masking |
| `src/Report.psm1` | Console summary and CSV |
| `tests/Run-Tests.ps1` | Tests with invented data and fake PnP cmdlets (no network) |
| `tests/Lint.ps1` | Windows PowerShell 5.1 syntax check (PSScriptAnalyzer) |

## Develop

```bash
pwsh -NoProfile -File tests/Run-Tests.ps1
```

```bash
pwsh -NoProfile -File tests/Lint.ps1
```

Pass = `TESTS ... failed=0` and `LINT ... findings=0`. GitHub Actions runs both on every push and pull request. Test data is invented; PII-shaped samples are built from parts at run time so no complete PII-shaped string is stored in this public repo.
