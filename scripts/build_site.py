#!/usr/bin/env python3
"""Build the shared parts of the site page (index.html) from README.md, so the README is the one source of truth.

Why: the README and the site said the same facts in two hand-written copies, and they drifted (the PAM, 2026-11-20:
"one needs to be source of truth"; the PAM chose the README). Run target: any machine with Python 3 (a build tool,
never run on a gov device).
What it does: README.md holds four marked sections, "<!-- site:NAME -->" to "<!-- /site:NAME -->". Each one is
turned into HTML and written into index.html between the generated-block marks
"<!-- >>> site:NAME: generated from README.md by scripts/build_site.py --write (do not edit) -->" and
"<!-- <<< site:NAME -->" (the kit's code standard C5 form). It also checks that the tool's breach text
(src/Report.psm1, $script:BreachAdvice) says the README's breach section, word for word.
  intro      one paragraph                               -> the paragraph under the tool's name
  key-facts  a table "| **Label:** | text |", then one line -> the aligned list of key facts, then that line
  looks-for  a table Type | Found when | Shown as | Confidence -> the table's rows (the first three columns)
  breach     the advice, a blank line, its citation      -> the breach note
Inputs:  README.md, index.html, src/Report.psm1.
Outputs: with --write, index.html; then, or with --check alone, "SITE PASS", or one line per finding and "SITE FAIL".
Usage:   python3 scripts/build_site.py --write    rebuild the site's shared parts, then check
         python3 scripts/build_site.py --check    check only (scripts/check.local.sh runs it)
         python3 scripts/build_site.py --test     self-test; prints SITE TEST PASS
Network: none.
Writes: index.html with --write only (the self-test: temporary files).
Credentials: none.
"""

import argparse
import html
import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
README = 'README.md'
SITE = 'index.html'
TOOL = 'src/Report.psm1'
COMMAND = 'scripts/build_site.py --write'
# The four shared sections, in README order (the PAM, 2026-11-20, README Source Preview v3).
SECTIONS = ('intro', 'key-facts', 'looks-for', 'breach')
# The site shows the first three columns of the README's "What it looks for" table (Confidence stays README only);
# their names must be these, as the site's hand-written table head says them (index.html <thead>).
SITE_HEAD = ('Type', 'Found when', 'Shown as')
SITE_COLUMNS = len(SITE_HEAD)
# Where the tool keeps its breach text: one PowerShell string, written as single-quoted parts joined with "+".
TOOL_VAR = re.compile(r"\$script:BreachAdvice\s*=\s*((?:'(?:[^']|'')*'\s*\+?\s*)+)")
TOOL_PART = re.compile(r"'((?:[^']|'')*)'")
# Markdown that GitHub renders but this builder does not: a site section that holds any of it fails the check, so the
# site never shows text that the README shows differently. An underscore inside a word is plain text on GitHub
# (AFI33-332_DAFGM2026-01), so only an underscore at a word's edge is refused.
UNSUPPORTED = ((r'[*`\\]', 'a lone * or ` or \\'), (r'(?<![A-Za-z0-9])_|_(?![A-Za-z0-9])', 'an _emphasis_ mark'),
               (r'~~', 'strikethrough'), (r'[\[\]<>]', 'a link, an HTML tag or a quote mark'),
               (r'&(#\d+|#x[0-9A-Fa-f]+|[A-Za-z]+);', 'an HTML entity'), (r'https?://|www\.', 'a bare web address'))
# A line that starts a list, a heading or a quote in Markdown.
BLOCK_START = r'\s*([-+*]\s|#|>|\d+[.)]\s)'
# A table's delimiter row (GFM): cells of dashes with optional alignment colons.
DELIMITER_ROW = r'\|(\s*:?-+:?\s*\|)+'
# HTML indents inside the generated blocks, matching the hand-written page around them.
INDENT_BLOCK = '        '
INDENT_ROW = '              '


def readme_section(text, name):
    """Return the text between a README section's marks, stripped; raise ValueError when it is missing.

    Args:
        text: the whole README.md.
        name: the section name, for example "key-facts".
    Returns:
        The section's Markdown, without its marks.
    """
    match = re.search(rf'<!-- site:{re.escape(name)} -->\n(.*?)\n<!-- /site:{re.escape(name)} -->', text, re.S)
    if not match:
        raise ValueError(f'{README}: no section <!-- site:{name} --> ... <!-- /site:{name} -->')
    return match.group(1).strip()


def refuse_marks(markdown):
    """Raise ValueError when plain text (outside `code`, **bold** marks removed) holds Markdown GitHub would render.

    Args:
        markdown: one line of text, without its `code` spans.
    """
    text = re.sub(r'\*\*([^*`]+)\*\*', r'\1', markdown)
    for pattern, what in UNSUPPORTED:
        if re.search(pattern, text):
            raise ValueError(f'{README}: unsupported Markdown ({what}) in a site section; only **bold** and `code`: {markdown}')


def inline(markdown):
    """Turn one line of Markdown (text, **bold**, `code`) into escaped HTML; raise ValueError on any other Markdown.

    Args:
        markdown: one line of text; text inside `code` is kept as is and never made bold.
    Returns:
        The escaped HTML.
    """
    out = []
    for piece in re.split(r'(`[^`]+`)', markdown):
        if piece.startswith('`') and piece.endswith('`') and len(piece) > 1:
            out.append('<code>' + html.escape(piece[1:-1], quote=False) + '</code>')
            continue
        refuse_marks(piece)
        out.append(re.sub(r'\*\*([^*`]+)\*\*', r'<strong>\1</strong>', html.escape(piece, quote=False)))
    return ''.join(out)


def paragraphs(markdown):
    """Return the paragraphs of plain Markdown text (blank lines split them; lines inside one are joined).

    Args:
        markdown: text without a table.
    Returns:
        The paragraphs, each on one line; raise ValueError when a line starts a list, heading or quote.
    """
    for line in markdown.splitlines():
        if re.match(BLOCK_START, line):
            raise ValueError(f'{README}: a list, heading or quote is not supported in a site section: {line}')
    return [' '.join(part.split()) for part in re.split(r'\n\s*\n', markdown) if part.strip()]


def split_row(line):
    """Split one Markdown table row into cell texts; raise ValueError on an escaped pipe.

    Args:
        line: one table row, for example "| SSN | 3-2-4 digits |". GitHub splits on every "|", even inside `code`.
    Returns:
        The cell texts, stripped.
    """
    if '\\|' in line:
        raise ValueError(f'{README}: an escaped "|" is not supported in a site table: {line}')
    return [cell.strip() for cell in line.strip()[1:-1].split('|')]


def table(markdown):
    """Return a section's table (header cells and body rows) and the text after it.

    Args:
        markdown: the section: a table first (every line starts and ends with "|", the second line is the
            delimiter row), then optional plain text after a blank line.
    Returns:
        A triple: the header cells, the body rows as lists of cell texts, and the text after the table.
    Raises:
        ValueError: when the table is missing, malformed, or a row's cell count differs from the header's.
    """
    lines = markdown.splitlines()
    end = next((n for n, line in enumerate(lines) if not line.strip().startswith('|')), len(lines))
    rows_text, rest = lines[:end], '\n'.join(lines[end:])
    if len(rows_text) < 3 or any(not line.strip().endswith('|') for line in rows_text):
        raise ValueError(f'{README}: a site section must start with a table: every row starts and ends with "|"')
    if not re.fullmatch(DELIMITER_ROW, rows_text[1].strip()):
        raise ValueError(f'{README}: the second table line must be the delimiter row, such as |---|---|')
    if '|' in rest or (rest.strip() and not rest.startswith('\n')):
        raise ValueError(f'{README}: text after a site table needs a blank line, and no other table rows')
    head, rows = split_row(rows_text[0]), [split_row(line) for line in rows_text[2:]]
    for row in rows:
        if len(row) != len(head):
            raise ValueError(f'{README}: a table row has {len(row)} cells, its header {len(head)}: {row}')
    return head, rows, rest


def build_intro(markdown):
    """Return the site's intro paragraph."""
    parts = paragraphs(markdown)
    if len(parts) != 1:
        raise ValueError(f'{README}: site:intro must be one paragraph')
    return f'{INDENT_BLOCK}<p>{inline(parts[0])}</p>'


def build_key_facts(markdown):
    """Return the site's aligned key-facts list and the line after it.

    Args:
        markdown: the key-facts section: a two-column table, a blank line, one closing line.
    Returns:
        The HTML: a definition list (label, text) and a small paragraph.
    """
    head, rows, rest = table(markdown)
    if len(head) != 2 or any(not re.fullmatch(r'\*\*[^*]+:\*\*', row[0]) for row in rows):
        raise ValueError(f'{README}: site:key-facts rows must be "| **Label:** | text |"')
    items = ''.join(f'\n{INDENT_BLOCK}  <dt>{inline(row[0][2:-2])}</dt><dd>{inline(row[1])}</dd>' for row in rows)
    tail = ''.join(f'\n{INDENT_BLOCK}<p class="small mut">{inline(part)}</p>' for part in paragraphs(rest))
    return f'{INDENT_BLOCK}<dl class="facts">{items}\n{INDENT_BLOCK}</dl>{tail}'


def build_looks_for(markdown):
    """Return the table rows of "What it looks for" (Type, Found when, Shown as).

    Args:
        markdown: the looks-for section: a table with at least three columns.
    Returns:
        The HTML rows, one per line.
    """
    head, rows, rest = table(markdown)
    if rest.strip():
        raise ValueError(f'{README}: site:looks-for holds only its table')
    if tuple(head[:SITE_COLUMNS]) != SITE_HEAD:
        raise ValueError(f'{README}: site:looks-for must start with the columns {" | ".join(SITE_HEAD)} (the site\'s table head)')
    return '\n'.join(INDENT_ROW + '<tr>' + ''.join(f'<td>{inline(cell)}</td>' for cell in row[:SITE_COLUMNS]) + '</tr>'
                     for row in rows)


def breach_parts(markdown):
    """Return the breach advice and its citation; raise ValueError unless the section has exactly these two paragraphs."""
    parts = paragraphs(markdown)
    if len(parts) != 2:
        raise ValueError(f'{README}: site:breach must hold the advice, a blank line, then its citation')
    return parts[0], parts[1]


def build_breach(markdown):
    """Return the site's breach note: the advice, then its citation in small text."""
    advice, citation = breach_parts(markdown)
    return f'{INDENT_BLOCK}<p class="note">{inline(advice)}<br><span class="mut small">{inline(citation)}</span></p>'


BUILDERS = {'intro': build_intro, 'key-facts': build_key_facts, 'looks-for': build_looks_for, 'breach': build_breach}


def tool_text(advice, citation):
    """Return the breach text the tool must show: the advice with its citation in brackets."""
    return f'{advice.removesuffix(".")} ({citation.removesuffix(".")}).'


def block_pattern(name):
    """Return the regex of a generated block in index.html: its two marks and what lies between them."""
    start = f'<!-- >>> site:{name}: generated from {README} by {COMMAND} (do not edit) -->'
    return re.compile(rf'({re.escape(start)}\n)(.*?)(\n[ \t]*<!-- <<< site:{re.escape(name)} -->)', re.S)


def render(readme, site):
    """Return the site page with every shared block rebuilt from the README.

    Args:
        readme: the text of README.md.
        site: the text of index.html.
    Returns:
        The new text of index.html.
    Raises:
        ValueError: when a README section or a site block is missing or has the wrong form.
    """
    for name in SECTIONS:
        pattern = block_pattern(name)
        if len(pattern.findall(site)) != 1:
            raise ValueError(f'{SITE}: needs exactly one generated block for site:{name}')
        body = BUILDERS[name](readme_section(readme, name))
        site = pattern.sub(lambda match, body=body: match.group(1) + body + match.group(3), site)
    return site


def check_tool(readme, tool):
    """Return findings when the tool's breach text differs from the README's breach section.

    Args:
        readme: the text of README.md.
        tool: the text of src/Report.psm1.
    Returns:
        A list of finding lines, empty when they match.
    """
    match = TOOL_VAR.search(tool)
    if not match:
        return [f'{TOOL}: no $script:BreachAdvice string found']
    said = ''.join(part.replace("''", "'") for part in TOOL_PART.findall(match.group(1)))
    wanted = tool_text(*breach_parts(readme_section(readme, 'breach')))
    return [] if said == wanted else [f'{TOOL}: $script:BreachAdvice differs from {README} site:breach; it must say: {wanted}']


def run(root, write):
    """Check (and with write, rebuild) the site under root; return the list of findings.

    Args:
        root: the repository folder.
        write: True to rewrite index.html first.
    Returns:
        The findings, empty when everything passes.
    """
    try:
        readme = (root / README).read_text(encoding='utf-8')
        site = (root / SITE).read_text(encoding='utf-8')
        fresh = render(readme, site)
        findings = check_tool(readme, (root / TOOL).read_text(encoding='utf-8'))
    except (OSError, ValueError) as err:
        return [str(err)]
    if write and fresh != site:
        (root / SITE).write_text(fresh, encoding='utf-8')
        site = fresh
    if fresh != site:
        findings.append(f'{SITE}: out of date with {README}; run: python3 {COMMAND}')
    return findings


def self_test():
    """Run the build and check on invented files; return the exit code (0 when every case passes).

    Returns:
        0 when every case passes, else 1.
    """
    readme = ('<!-- site:intro -->\nA safe_tool & more.\n<!-- /site:intro -->\n'
              '<!-- site:key-facts -->\n| Fact | |\n|---|---|\n| **Read-only:** | it changes `nothing`. |\n\nFirst line.\nAsk & learn.\n'
              '<!-- /site:key-facts -->\n'
              '<!-- site:looks-for -->\n| Type | Found when | Shown as | Confidence |\n|---|---|---|---|\n'
              '| SSN | 3-2-4 digits | `***-**-1234` | High |\n<!-- /site:looks-for -->\n'
              '<!-- site:breach -->\nReport it now.\n\nRule para 1.\n<!-- /site:breach -->\n')
    marks = ''.join(f'<!-- >>> site:{name}: generated from {README} by {COMMAND} (do not edit) -->\nold\n'
                    f'<!-- <<< site:{name} -->\n' for name in SECTIONS)
    tool = "$script:BreachAdvice = 'Report it now ' +\n    '(Rule para 1).'\n"
    cases = [('a fresh build passes', None, [], True),
             ('a stale site fails', 'stale', ['out of date'], False),
             ('a tool text that differs fails', 'tool', ['BreachAdvice differs'], True),
             ('a missing README section fails', 'section', ['no section'], True),
             ('a key fact without a label fails', 'label', ['**Label:**'], True),
             ('a renamed site column fails', 'column', ['must start with the columns'], True),
             ('a row with a wrong cell count fails', 'cells', ['cells, its header'], True),
             ('unsupported Markdown fails', 'italic', ['unsupported Markdown'], True),
             ('a link fails', 'link', ['unsupported Markdown (a link'], True),
             ('a "|" inside code fails, as GitHub splits the cell there', 'pipe', ['cells, its header'], True),
             ('a table without its delimiter row fails', 'delimiter', ['delimiter row'], True),
             ('a list in a section fails', 'list', ['list, heading or quote'], True)]
    failed = 0
    for label, broken, expect, write in cases:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'src').mkdir()
            text = readme.replace('<!-- site:intro -->', '<!-- site:gone -->') if broken == 'section' else readme
            text = text.replace('**Read-only:**', 'Read-only') if broken == 'label' else text
            text = text.replace('| Shown as |', '| Shown |') if broken == 'column' else text
            text = text.replace('| High |', '| High | extra |') if broken == 'cells' else text
            text = text.replace('Report it now.', 'Report it *now*.') if broken == 'italic' else text
            text = text.replace('Rule para 1.', 'Rule [para 1](x).') if broken == 'link' else text
            text = text.replace('| 3-2-4 digits |', '| 3-2-4 `a|b` digits |') if broken == 'pipe' else text
            text = text.replace('|---|---|---|---|\n', '| a | b | c | d |\n') if broken == 'delimiter' else text
            text = text.replace('First line.', '- First line.') if broken == 'list' else text
            (root / README).write_text(text, encoding='utf-8')
            (root / SITE).write_text(marks, encoding='utf-8')
            (root / TOOL).write_text(tool.replace('now', 'later') if broken == 'tool' else tool, encoding='utf-8')
            found = run(root, write)
            ok = (not found) if not expect else all(any(e in line for line in found) for e in expect)
            if label == 'a fresh build passes':
                built = (root / SITE).read_text(encoding='utf-8')
                ok = ok and '<dt>Read-only:</dt><dd>it changes <code>nothing</code>.</dd>' in built and 'Ask &amp; learn.' in built
                ok = ok and '<td>3-2-4 digits</td><td><code>***-**-1234</code></td></tr>' in built
                ok = ok and '>High<' not in built and 'First line. Ask &amp; learn.</p>' in built
            failed += not ok
            print(('ok   ' if ok else 'FAIL ') + label + ('' if ok else f': {found}'))
    print('SITE TEST PASS' if not failed else 'SITE TEST FAIL')
    return 1 if failed else 0


def main():
    """Parse the arguments and run the rebuild, the check or the self-test.

    Returns:
        The exit code: 0 on pass, else 1.
    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--write', action='store_true', help='rebuild the shared parts of index.html, then check')
    mode.add_argument('--check', action='store_true', help='check only')
    mode.add_argument('--test', action='store_true', help='self-test')
    args = parser.parse_args()
    if args.test:
        return self_test()
    found = run(ROOT, args.write)
    for line in found:
        print(line)
    print('SITE PASS' if not found else 'SITE FAIL')
    return 1 if found else 0


if __name__ == '__main__':
    sys.exit(main())
