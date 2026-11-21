# Ion backlog item: who changed a file with PII, and when (2026-11-21)

Status: on the implementation list; requirement confirmed by the PAM 2026-11-21; plan not written yet. Source: the PAM's request, relayed by the RADE session on 2026-11-21 ("before a file is scanned, can you see who last modified it in case PII is found within and someone must be held accountable?" and "when it was last modified so the duration of the potential incident can be determined").

## What it would add
- For each file with a hit: Modified and Modified By (Editor), Created and Created By (Author), read-only with the person's own access (fields of the list item; no extra site, team or user is listed: "No reconnaissance").
- Where version history is readable: the first version that holds the PII, which is the start of the exposure; "last modified" is only its end. The last editor is not always the person who added the PII.
- Cost: a version-history read per file with a hit (not per file scanned); network drives give only owner and dates.

## Decision (the PAM, 2026-11-21, confirmed in this chat: docs/DECISIONS.md)
- Teams/SharePoint: who put the PII in and when (version history), plus Created/Created By and Modified/Modified By.
- Local and network folders: who created the file and who last modified it, and when.
- Each person by e-mail address; the report is CUI and stays inside DoD365. This repo holds only invented addresses.

## What Windows records for local and network files (to check on the gov device)
- Created and last-modified time: always. The file owner (ACL): usually the creator, as DOMAIN\user, not an e-mail address.
- Who last modified: Windows does not store it on the file. Office files store it in the file (docProps/core.xml: creator,
  lastModifiedBy), as the display name the author's Office used. Other file types have no last-modified-by, and the
  report says so.
- An e-mail address for a DOMAIN\user or a display name needs a directory lookup of that one person.
