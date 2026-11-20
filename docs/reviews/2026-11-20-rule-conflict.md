# Rule conflict check: 2026-11-20

Changed lines (AGENTS.md "Project rules"; a project rule, not marked "cannot be changed", so the short record applies):

- Before: Public repo (2026-11-20): never commit a wing's name or location, a real site path (only `https://usaf.dps.mil/sites/<team-site>`), DODIDs or gov-network details.
- After: Public repo (2026-11-20): never commit a wing's name or location, a real site path (only the placeholders `https://usaf.dps.mil/sites/<team-site>` and `https://usaf.dps.mil/teams/<team-site>`), DODIDs or gov-network details.

Approval: the PAM, 2026-11-20, on the exact "After" text shown in the session's turn summary (SUGGESTED).

Step 1 (search AGENTS.md and docs/DECISIONS.md for the same topic):

- AGENTS.md: no other line names a site path placeholder; "Data boundary" (no site paths in the repo) is kept, since both forms are placeholders, not real paths.
- docs/DECISIONS.md: the `.check` row of 2026-11-20 says "The tool's docs use `https://usaf.dps.mil/sites/<team-site>`". The new row of 2026-11-20 (Public repo rule) and the PR #6 row record the `/teams/` placeholder; the newer rows win. No conflict left open.

rule-conflict-check: c28acbb7f4d4
