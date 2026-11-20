# Decisions: 33-332.Z

A record, read on demand: a decision binds the task or thing it names. A decision that must bind every session is a line in AGENTS.md "Project rules". Newest first. Why the kit defaults are what they are: https://github.com/rhope-ops/archimedes_lever/blob/main/docs/DECISIONS.md. "By" = the owner for a decision. "By" = Agent for a recommendation that the owner has not approved yet. Times are America/Chicago.

| Date | Decision | By | Replaces |
|---|---|---|---|
| 2026-11-20 | Higher level of checks for DoD/M365 code: `.check` with `profile: dod`, `code-c1-safety: on`, `code-c2: full`, `code-max-line-chars: 200` and `code-generated: paste/*.ps1`; the code-standard gate of 90-161.Z/RADE; the PowerShell tests on Windows PowerShell 5.1 in CI (ps51 workflow). Scope: this repo. | the PFM | — |
| 2026-11-20 | Kit sync: rhope-ops/archimedes_lever 6.0.0 (commit 763b163; no v6.0.0 tag yet), tier full, so the required check `checks` (the kit's content check) runs. Scope: this repo. | the PFM | — |
| 2026-11-20 | Branch rule: the required status check is `checks` (was `scan`; the kit's job is shown as `checks`). Scope: this repo. | the PFM | — |
| 2026-11-20 | Gov-device use: the tool is one generated text pasted into PowerShell ISE, as in RADE (ISSM ruling 2026-11-13: scripts without elevated credentials may be pasted into ISE and run); no download, no Unblock-File, no execution-policy change. Scope: the tools of this repo and their instructions. | the PFM | — |
| 2026-11-20 | Logo: γ Gamma, "L2 Lens, pages" (Logo Studies, standard v2.9) for 33-332.Z; the site header shows it on the left and the Core mark on the right at the same size. Scope: the site. | the PFM | — |
| 2026-11-20 | Teams PII Spot Check: every found value masked wherever it is shown (screen, progress bar, .txt, .csv, names, links, errors); report choice Both/Text/Csv; home address and labelled-name detectors; test samples built from parts at run time so no complete PII-shaped string is in the repo. Scope: the tool. | the PFM | — |
| 2026-11-20 | Public repo (was planned private; renamed from 314). Scope: this repo. | the PFM | — |
