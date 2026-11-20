# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Kit sync PR (this one): merge after the owner approves (lane C: instruction and guard files).
2. Then the PII tool as a fresh one-commit PR from main (the owner, 2026-11-20): the final code of PR #1
   (branch claude/ecstatic-johnson-c1mg5q, head 903c43b; six fresh reviews, no MAJOR), plus the slow-regex fix
   (atomic top-level label), docs using https://usaf.dps.mil/sites/<team-site>, and without its own
   tests/lint workflow (the `checks` job runs them). Close PR #1 with a link to it.
3. Then one PR: the paste bundle for PowerShell ISE (build script + generated paste/*.ps1, RADE style), README run
   steps without Unblock-File or execution policy, and the site page (index.html): header like 90-161.Z (γ logo,
   name, subtitle "Privacy Act Management", line "Simple tools to help you protect PII.", Core logo right, same size),
   FULCRUM colours only (no teal), much less on the page, the builder outputs the paste text. Preview first.
4. The owner: add the secret PII_GUARD_DENYLIST to this repo (Settings > Secrets and variables > Actions), as in 90-161.Z;
   the DoD profile fails the content check without it. Turn on GitHub Pages (main, / root) once index.html merges.
