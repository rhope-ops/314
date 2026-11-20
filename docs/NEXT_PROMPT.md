# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Done 2026-11-20: kit sync (PR #2), the PII tool and its paste bundle (PR #3), the site page (PR #4),
   PII Scan review follow-ups (PR #5).
2. Done 2026-11-20: PII Scan takes a Teams link or a SharePoint address (PR #6, merged 084a7ba). Preview:
   https://claude.ai/artifact/4vYLihr6qHtXFMtkyRQ4Ub
3. Done 2026-11-20: PR #5 review minors, Key Facts and hint wording, Public repo rule with the /teams/
   placeholder (PR #7, merged 1b79fa8).
4. Done 2026-11-20 (PR #9, merged 5d50c64): the tenant-root lookup that hung a Teams-link run is removed (a Teams
   link asks for the site address, then checks its group ID); no choices (all channels, posts always tried, both
   reports); progress bar and per-channel counts (src/ScanMeter.psm1); page without Configuration, run steps at
   the top of the script; for a team member, not only an owner. The PAM: test one run on DoD365.
5. Open 2026-11-20 (PR #11): Ion (DoD365 PII Reporter): Excel read cell by cell (false positives), per-file
   progress bar, top-of-library files, posts off, one .xlsx workbook (Summary first, a tab per type worst first,
   Low confidence, Not scanned) in Documents\Ion\<site>\<date>. The PAM: approve the preview, then test a run.
6. GitHub Pages is on (live copy checked 2026-11-20). Denylist secret: added 2026-11-20.
7. Known gap (MINOR, safe): pasted text with a second address after a Teams link without a groupId can supply
   that later groupId; the site-group check still holds.
