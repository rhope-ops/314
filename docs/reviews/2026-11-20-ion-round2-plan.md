# Ion round 2 (plan, approved by the PAM 2026-11-20 in the question dialog)

Done = all steps built, `scripts/check.sh` CHECK PASS, fresh review without MAJOR, PR merged, paste bundle rebuilt.

1. Folder rows: folder name first, in one fixed-width column, teal (DarkCyan); labels gray, numbers white; possible
   PII above 0 red, not scanned above 0 yellow. The live total row uses the same column ("Total"), so all rows line up.
2. Libraries: scan every visible document library (BaseTemplate 101, not hidden) of the one site given; skip Style
   Library and Form Templates; no subsites. Top-level groups are "Library / Folder" when the site has more than one library.
3. PDF: a basic built-in reader (no module): fillable-form field values, and text from Flate-compressed content streams
   (Tj/TJ, with ToUnicode maps where present). No text found = "No text layer (scanned image?)"; encrypted = its own reason.
4. Teams link: the PAM first runs a read-only test snippet (host-root sign-in, one-group GetSiteStatus lookup). Build it
   into Ion only if it works (separate change).
