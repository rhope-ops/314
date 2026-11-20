# Teams channel PII spot check

A read-only PowerShell script that a **team owner** runs to find files (and, if allowed, posts) in one Microsoft Teams team that may hold PII. It shows a readable summary on screen and saves a readable text report, a CSV, or both (you choose). **Every found value is masked** (for example `***-**-1234`), so the report is not a PII spill itself.

> This is a spot check, not a compliance control. The standard tool for PII in Teams is **Microsoft Purview** (DLP policies and Content Search), run by your wing's privacy or cybersecurity office.

## What it scans

| Content | Scanned? | How |
|---|---|---|
| Files in standard channels (Shared tab, formerly Files) | Yes | Your own CAC sign-in to the team's SharePoint site |
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

Needs: Windows PowerShell 5.1 and PnP.PowerShell 1.12.0. No other modules, no admin rights, nothing to download or unblock. You paste one text into PowerShell ISE: pasted text is not a script file, so the execution policy does not apply (ISSM ruling 2026-11-13).

1. Open the paste text: [paste/33332Z-pii-spot-check.ps1](https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/paste/33332Z-pii-spot-check.ps1). Select all (Ctrl+A) and copy (Ctrl+C).
2. Open **Windows PowerShell ISE**, choose **File > New**, and paste into the script pane (Ctrl+V).
3. Optional: fill in the **Settings** lines at the top (Teams link or site address, channel, posts, report). Anything left empty is asked for when it runs.
4. Press **F5**. It first asks for anything you left empty in the Settings: the team (see "The team to scan" below) and what to save (**1** Both, press Enter; **2** Text; **3** CSV).
5. Sign in with your CAC in the window that opens (a Teams link opens two: one to find the team's site, one for the site), then pick the channels: a number, several numbers such as `1,3`, or `A` for all.
6. Read the summary on screen. The saved file paths are printed at the end (default folder: `Documents\PII-Scan`).

The team to scan, one of:

- **Teams link:** in Teams, open the team or one of its channels, click the **...** next to its name, then **Copy link** (`https://<teams-host>/l/team/...?groupId=...`). The script asks SharePoint at `https://usaf.dps.mil` for the site of that one team (by the link's group ID), then checks that the site belongs to that team before it reads anything. If the lookup does not work, it asks for the site address.
- **SharePoint address:** in the team's **Shared** tab, click **...** > **Open in SharePoint** and copy the address (`https://usaf.dps.mil/teams/<team-site>`; some sites use `/sites/`). A file's **Copy link** from the **Shared** tab works too. Anything after the site name is cut off. `https://` may be left out.

What each report holds:

| Report | Use it to | Holds |
|---|---|---|
| Screen and `.txt` | Read and print | Totals; items with possible PII (highest confidence first); hits per type; each hit with its masked value and masked context; items not scanned |
| `.csv` | Sort and filter in Excel | One row per hit or note: Source, Channel, Item, Detector, Confidence, MaskedValue, Context, Modified, Link |

Settings at the top of the paste text:

| Setting | Meaning |
|---|---|
| `$SiteUrl` | The team's Teams link, or its site `https://usaf.dps.mil/teams/<team-site>`; empty = asked |
| `$Channel` | A channel folder name, or `All`; empty = pick from a list |
| `$IncludePosts` | `$true` = also try posts and replies (a second sign-in; may be BLOCKED) |
| `$Report` | `Both`, `Text` or `Csv`; empty = asked |
| `$OutputFolder` | Where to save the report files (default `Documents\PII-Scan`) |
| `$MaxFileMB` | Size limit per file (default 25) |

## Safety

- **Read-only:** it never changes, moves or deletes anything in Teams or SharePoint.
- **One team only:** it touches only the team you give: its site, and for a Teams link one lookup of that team's site. It never lists other sites or teams.
- **No stored credentials:** your own CAC sign-in each run.
- **Masked report:** the screen, the `.txt` and the `.csv` never hold a full value. The context snippet is cut from text that is already masked (digits become `#`, found values and e-mail names become `*`), so a cut cannot expose part of a value. Names the report shows as they are (file, folder and channel names, error messages) get a stricter mask: found values are masked, then every digit becomes `#` and every e-mail name `***` (so `FY27 Plan.docx` shows as `FY## Plan.docx`). A link goes to the file only when its path has nothing to mask; else to the deepest clean folder.
- **Safe CSV:** a cell that Excel would read as a formula gets a leading `'`.
- **Stops cleanly:** if something fails part way or you press Ctrl+C, the report is still saved and says INCOMPLETE; the script signs out even if the report cannot be saved.
- **Handle the reports as CUI anyway:** file names and masked fragments still point to where PII is.
- PII where it should not be is a potential or actual breach; report it to your servicing Privacy Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, 7 July 2026).

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
| `tests/*.Tests.ps1` | One test file per module, a start-to-end run, and a run of the paste text |
| `tests/Lint.ps1` | Windows PowerShell 5.1 syntax check (PSScriptAnalyzer) |
| `paste/33332Z-pii-spot-check.ps1` | GENERATED: the one text you paste into ISE (the five modules and the script, with a Settings block on top) |
| `scripts/build_paste.py` | Builds the paste text from `src/`; `--check` fails when it is out of date |
| `ref/afi33-332.pdf`, `ref/afi33-332.txt` | The rule source: AFI 33-332 with AFI33-332_DAFGM2026-01 (edition: docs/DECISIONS.md) |

## Standards and policy it follows

| Source | What it means here |
|---|---|
| [docs/code-standard-33-332.md](docs/code-standard-33-332.md): kit rules C1–C11 with `code-c1-safety: on`, `code-c2: full`, `code-max-line-chars: 200`, and the RADE extras X1–X7 | Headers with safety lines and run target; full help for each function; functions ≤ 60 lines, files ≤ 600, lines ≤ 200; named values with their source; every discarded error says why; `Set-StrictMode -Version 2`. Checked on every pull request by `scripts/check.sh` (the `checks` job) |
| Windows PowerShell 5.1, PnP.PowerShell 1.12.0, no other modules | Lint checks 5.1 syntax; the `ps51` job runs the tests on Windows PowerShell 5.1; only .NET built-ins are used |
| Paste and run (ISSM ruling 2026-11-13) | One generated text pasted into ISE; no download, no unblocking, no execution-policy change |
| No network reconnaissance | Only the one team given (its site; for a Teams link, one lookup of that one group's site); nothing lists other sites or teams |
| Government data stays in DoD365 | File content is read in memory only; reports stay on your computer; the tests use invented data only |
| Credentials | Your own CAC sign-in each run; nothing stored |
| No scripts in results | Script files are not read; only their masked names can appear in a report |
| Decisions in comments | Each names who decided and the date (for example "the PAM, 2026-11-20") |

## Develop

```bash
scripts/check.sh
```

Pass = last line `CHECK PASS`. It runs the kit checks, the content check and `scripts/check.local.sh`: the paste-bundle check, the PowerShell tests (`tests/Run-Tests.ps1`: invented data, fake PnP cmdlets, no network; the last test runs the paste bundle itself) and the lint (`tests/Lint.ps1`, Windows PowerShell 5.1 syntax). After a change in `src/`, rebuild the paste text:

```bash
python3 scripts/build_paste.py
```

On pull requests, the `checks` job runs the same, and the `ps51` job runs the tests on Windows PowerShell 5.1. PSScriptAnalyzerSettings.psd1 says which style rules are off and why. Test data is invented; PII-shaped samples are built from parts at run time, so no complete PII-shaped string is stored in this public repo.
