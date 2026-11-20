# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Done 2026-11-20: kit sync (PR #2), the PII tool and its paste bundle (PR #3), the site page (PR #4),
   PII Scan review follow-ups (PR #5).
2. Done 2026-11-20: PII Scan takes a Teams link or a SharePoint address (PR #6, merged 084a7ba). Preview:
   https://claude.ai/artifact/4vYLihr6qHtXFMtkyRQ4Ub
3. Done 2026-11-20: PR #5 review minors, Key Facts and hint wording, Public repo rule with the /teams/
   placeholder (PR #7, merged 1b79fa8).
4. 2026-11-20: the first Teams-link run on DoD365 hung at the tenant-root lookup; the lookup is removed (a Teams
   link now asks for the site address, then checks its group ID). The PAM: test one run again.
5. The PAM: turn on GitHub Pages (main, / root) if not on yet. (Denylist secret: added 2026-11-20.)
6. Known gap (MINOR, safe): pasted text with a second address after a Teams link without a groupId can supply
   that later groupId; the site-group check still holds.
