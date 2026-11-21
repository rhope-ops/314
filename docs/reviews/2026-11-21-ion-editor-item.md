# Ion backlog item: who changed a file with PII, and when (2026-11-21)

Status: on the implementation list; not planned or approved yet. Source: the PAM's request, relayed by the RADE session on 2026-11-21 ("before a file is scanned, can you see who last modified it in case PII is found within and someone must be held accountable?" and "when it was last modified so the duration of the potential incident can be determined").

## What it would add
- For each file with a hit: Modified and Modified By (Editor), Created and Created By (Author), read-only with the person's own access (fields of the list item; no extra site, team or user is listed: "No reconnaissance").
- Where version history is readable: the first version that holds the PII, which is the start of the exposure; "last modified" is only its end. The last editor is not always the person who added the PII.
- Cost: a version-history read per file with a hit (not per file scanned); network drives give only owner and dates.

## Relayed decision (to be confirmed by the PAM in their own message, AGENTS.md "Owner changes")
- 2026-11-21, relayed by the RADE session: the report shows each editor by e-mail address; the report is CUI and stays inside DoD365. Scope: the Ion report. Once confirmed: a docs/DECISIONS.md row. This repo still holds only invented addresses (Data boundary).
