# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Kit sync PR (this one): merge after the PFM approves (lane C: instruction and guard files).
2. Then PR #1 (Teams PII Spot Check): merge main into its branch so the `checks` and `ps51` jobs run; merge when green
   (its fresh reviews are done: no MAJOR). Carry over: the slow-regex MINOR (make the top-level label atomic) and
   remove its own tests/lint workflow (the `checks` job runs them through scripts/check.local.sh).
3. Then one PR: the paste bundle for PowerShell ISE (build script + generated paste/*.ps1, RADE style), README run
   steps without Unblock-File or execution policy, and the site page (index.html): header like 90-161.Z (γ logo,
   name, subtitle "Privacy Act Management", line "Simple tools to help you protect PII.", Core logo right, same size),
   FULCRUM colours only (no teal), much less on the page, the builder outputs the paste text. Preview first.
4. The PFM: add the secret PII_GUARD_DENYLIST to this repo (Settings > Secrets and variables > Actions), as in 90-161.Z;
   the DoD profile fails the content check without it. Turn on GitHub Pages (main, / root) once index.html merges.
