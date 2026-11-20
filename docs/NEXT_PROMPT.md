# Next session: 33-332.Z

Title: 33-332.Z · <task> · <date>. Read AGENTS.md first.

1. Done 2026-11-20: kit sync (PR #2, merged a25b74d).
2. Done 2026-11-20: the PII tool with its paste bundle (PR #3, merged dec64b6); PR #1 closed.
3. The site page (index.html, branch claude/site-page): the PAM's 8 preview comments applied (preview v4:
   https://claude.ai/artifact/R2QH26gZXYNjE6GUVxtJ4s); fresh review, then merge (the PAM said publish).
   Next: a follow-up PR with PR #3's review minors (bundle test isolation, $Report check, dotted-run speed,
   README step order, strict mode after a run in ISE); ask the PAM whether the report's breach line takes the
   page's wording.
4. The PAM: turn on GitHub Pages (main, / root) once index.html merges. (Denylist secret: added 2026-11-20.)
5. 2026-11-20: PII Scan takes a Teams link or a SharePoint address (PR #6,
   https://github.com/rhope-ops/33-332.Z/pull/6; preview https://claude.ai/artifact/4vYLihr6qHtXFMtkyRQ4Ub).
   Lane C (UI change): merge only after the PAM approves the preview. Open: the PAM tests the Teams-link lookup
   on DoD365 (not testable here); the PAM decides on a `.check` allow line for the Teams host so docs can show it.
6. Queued (the PAM, 2026-11-20): after PR #6 merges, one new PR with the PR #5 review minors: (a) Bundle test
   reaches the outer finally (strict mode off, error setting back, sign-out ran); (b) DECISIONS edition row takes
   $script:BreachAdvice wording + a dated row; (c) Remove-OverlappingHits linear (sort, sweep; test many overlaps);
   (d) cite the DAFGM2026-01 page-1 paragraph for "without delay"; (e) refresh this file.
