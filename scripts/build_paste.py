#!/usr/bin/env python3
"""Build the paste bundle of the Teams PII Spot Check: one PowerShell text that a person pastes into Windows
PowerShell ISE and runs with F5 (AGENTS.md "Platform": paste and run, never an execution-policy change).

Why: a gov computer runs pasted text in ISE, not downloaded files (ISSM ruling 2026-11-13). A pasted text loads no
module file and runs no script file, so the execution policy does not apply to it.
What it writes: paste/33332Z-pii-spot-check.ps1, in this order:
  1. a header (GENERATED, the sources with their SHA-256, how to run, Network/Writes/Credentials, a contents list);
  2. the Settings block: plain variables the person or the site page sets (site, channel, posts, report);
  3. the five modules of src/, inline (no Import-Module), in dependency order;
  4. the body of src/Find-ChannelPII.ps1 (from its "# ==== Body ====" line to the end);
  5. a Clean-up block: parts 3 and 4 run inside try/finally, which turns strict mode off and puts the ISE session's
     error setting back when the run ends.
Inputs:  src/PiiDetectors.psm1, src/PiiMasking.psm1, src/TextExtract.psm1, src/Report.psm1, src/ChannelScan.psm1,
         src/Find-ChannelPII.ps1.
Outputs: the bundle file; with --check, nothing: exit 1 when the committed bundle differs from a fresh build.
Usage:   python3 scripts/build_paste.py            build
         python3 scripts/build_paste.py --check    check (scripts/check.local.sh runs it)
Network: none.
Writes: paste/33332Z-pii-spot-check.ps1 (not with --check).
Credentials: none.
"""

import argparse
import hashlib
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / 'paste' / '33332Z-pii-spot-check.ps1'
MAIN = 'src/Find-ChannelPII.ps1'
# Dependency order: each module uses only those listed before it (their headers say so).
MODULES = ('src/PiiDetectors.psm1', 'src/PiiMasking.psm1', 'src/TextExtract.psm1', 'src/Report.psm1',
           'src/ChannelScan.psm1')
BODY_MARK = '# ==== Body ===='
# The end of the bundle: strict mode off and the session's error setting back, also after an error or Ctrl+C, so
# the ISE session the person keeps working in is as it was before F5.
CLEANUP = """# ==== Clean-up ====
} finally {
    Set-StrictMode -Off
    $ErrorActionPreference = $previousErrorAction
}
"""

# Module lines that do not belong in one pasted script: the export list (a module-only command, which may run on
# to the next lines) is dropped; Set-StrictMode is set once in the Settings block.
EXPORT_START = 'Export-ModuleMember'
STRICT_LINE = 'Set-StrictMode -Version 2.0'

# The Settings block. The site page fills these values in; a person may edit them in ISE before F5.
SETTINGS = """# ==== Settings ==== (edit the values between the quotes, or let the site page fill them in)
$SiteUrl = ''            # the team's site, https://usaf.dps.mil/sites/<team-site>; empty = asked
$Channel = ''            # 'All' or one channel name; empty = pick from a list
$IncludePosts = $false   # $true = also try posts and replies (a second sign-in; may be BLOCKED)
$Report = ''             # 'Both', 'Text' or 'Csv'; empty = asked
$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PII-Scan')   # where the report is saved
$MaxFileMB = 25          # files larger than this are reported as NOT SCANNED

# ISE runs pasted text in your session: keep its error setting so the Clean-up block at the end can put it back.
$previousErrorAction = $ErrorActionPreference
try {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
"""


def sha256(text):
    """Return the SHA-256 hex digest of a text (UTF-8)."""
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def read(rel):
    """Return the text of a repo file with Unix line ends."""
    return (ROOT / rel).read_text(encoding='utf-8').replace('\r\n', '\n')


def module_code(rel):
    """Return a module's code for the bundle: without its export list and its own Set-StrictMode line.

    Args:
        rel: the module path relative to the repo root.
    Returns:
        The code as a string ending in one line break.
    """
    lines = read(rel).split('\n')
    if STRICT_LINE not in lines:
        sys.exit('FAIL %s: no "%s" line (every module sets strict mode)' % (rel, STRICT_LINE))
    start = next((i for i, line in enumerate(lines) if line.startswith(EXPORT_START)), None)
    if start is None:
        sys.exit('FAIL %s: no %s line' % (rel, EXPORT_START))
    end = start + 1
    # The export list may run on to indented or plain continuation lines until the first blank line or the file end.
    while end < len(lines) and lines[end].strip() and not lines[end].startswith(('function ', '#', '$')):
        end += 1
    kept = [line for line in lines[:start] + lines[end:] if line != STRICT_LINE]
    return '\n'.join(kept).rstrip('\n') + '\n'


def body_code():
    """Return the body of the main script: from its "# ==== Body ====" line to the end."""
    text = read(MAIN)
    at = text.find(BODY_MARK)
    if at < 0:
        sys.exit('FAIL %s: no "%s" line' % (MAIN, BODY_MARK))
    return text[at:].rstrip('\n') + '\n'


def check_names(parts):
    """Stop when two modules define the same function or $script: variable (they share one scope in the bundle).

    Args:
        parts: (path, code) pairs of the modules.
    Returns:
        None.
    """
    seen = {}
    for rel, code in parts:
        names = set(re.findall(r'^function ([\w-]+)', code, re.M)) | set(re.findall(r'^\$script:(\w+) =', code, re.M))
        for name in names:
            if name in seen:
                sys.exit('FAIL %s and %s both define %s; rename one (the bundle has one scope)' % (seen[name], rel, name))
            seen[name] = rel


def header(sources, contents):
    """Return the bundle header: GENERATED, sources, how to run, safety lines and the contents list.

    Args:
        sources: (path, sha256) pairs of the source files.
        contents: (line number, title) pairs of the bundle's parts.
    Returns:
        The header text.
    """
    out = ['<#',
           '  33-332.Z Teams PII Spot Check: paste bundle.',
           '  GENERATED by scripts/build_paste.py from the files below; never edit it: change the sources and rebuild',
           '  with: python3 scripts/build_paste.py']
    out += ['    %s, sha256 %s' % (rel, digest[:12]) for rel, digest in sources]
    out += ['',
            '  How to run: open Windows PowerShell ISE, File > New, paste ALL of this text into the script pane, set the',
            '  Settings below (or let https://rhope-ops.github.io/33-332.Z/ fill them in), press F5, sign in with your CAC.',
            '  Read-only: it never changes, moves or deletes anything in Teams or SharePoint. Every value found is masked.',
            '  Run target: Windows PowerShell 5.1 with PnP.PowerShell 1.12.0; no admin rights, no file to unblock.',
            '  Network: only the one site address in the Settings, and with $IncludePosts Microsoft Graph for that team.',
            '  Writes: only the report files in $OutputFolder (default Documents\\PII-Scan).',
            '  Credentials: your own CAC sign-in in the PnP sign-in window; nothing is stored.',
            '',
            '  Contents (line numbers of this file):']
    out += ['    line %5d  %s' % (line, title) for line, title in contents]
    out += ['#>', '']
    return '\n'.join(out) + '\n'


def build():
    """Build the bundle text from the sources.

    Returns:
        The whole bundle text: header, Settings, the modules and the body.
    """
    modules = [(rel, module_code(rel)) for rel in MODULES]
    check_names(modules)
    parts = [('Settings', SETTINGS)]
    parts += [('Module %s' % Path(rel).name, '# ==== Module %s ====\n%s' % (Path(rel).name, code)) for rel, code in modules]
    parts.append(('Body of Find-ChannelPII.ps1', body_code()))
    parts.append(('Clean-up', CLEANUP))
    sources = [(rel, sha256(read(rel))) for rel in MODULES + (MAIN,)]
    # Two passes: the header's own length fixes the line numbers of the parts.
    contents = [(0, title) for title, _ in parts]
    for _ in range(2):
        head = header(sources, contents)
        line, contents = head.count('\n') + 1, []
        for title, text in parts:
            contents.append((line, title))
            line += text.count('\n') + 1
    return head + '\n'.join(text for _, text in parts)


def main():
    """Build the bundle, or with --check compare it with the committed file. Exit 1 on a difference."""
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--check', action='store_true', help='fail when paste/ is not the fresh build')
    args = ap.parse_args()
    text = build()
    if args.check:
        if not BUNDLE.is_file() or BUNDLE.read_text(encoding='utf-8') != text:
            print('PASTE FAIL: %s is not the fresh build; run python3 scripts/build_paste.py and commit it'
                  % BUNDLE.relative_to(ROOT))
            sys.exit(1)
        print('PASTE PASS: %s is up to date' % BUNDLE.relative_to(ROOT))
        return
    BUNDLE.parent.mkdir(parents=True, exist_ok=True)
    BUNDLE.write_text(text, encoding='utf-8', newline='\n')
    print('PASTE BUILT: %s (%d lines)' % (BUNDLE.relative_to(ROOT), text.count('\n')))


if __name__ == '__main__':
    main()
