# Ion network drives (plan, approved by the PAM 2026-11-21 in the question dialog)

Build after PR #14 (round 2) merges; before or with round 3 as the PAM says. Done = built, `scripts/check.sh`
CHECK PASS, fresh review without MAJOR, merged, bundle rebuilt.

1. Target: a pasted path `X:\<folder>` (mapped drive) or `\\<server>\<share>\<folder>` (UNC); the folder and all its
   subfolders; nothing above it, no other share or server. No sign-in: the person's Windows logon.
2. Groups: each top-level subfolder is one row (like a channel), files in the pasted folder are "(top of the folder)".
3. Reading: files opened read-only (FileAccess.Read, FileShare.ReadWrite); the same readers (Office, PDF, text);
   folders that cannot be opened are BLOCKED rows; long paths (> 260) through the \\?\ prefix; reparse points not followed.
4. Speed: a RunspacePool (no module) whose number of files in flight scales with the load: start at 2, add one while
   files/MB per second rise, CPU < ~85% and no read errors; drop on falling speed, slow reads or errors; 1 to 8.
   The Total row shows the worker count. SharePoint stays one file at a time (one CAC sign-in).
5. Report: the same workbook; the Link is a file:/// link to the file; names masked with Protect-PiiName; results
   folder named after the last folder of the path. Docs and tests use placeholders only (Public repo rule).
