import re, html, os
from pypdf import PdfReader

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(REPO_ROOT, "source-pdfs")
OUT_DIR = REPO_ROOT

def pdf_to_text(path):
    reader = PdfReader(path)
    return "\n\n".join(page.extract_text() or "" for page in reader.pages)

DOCS = [
    {"pdf": "2001-16.PDF", "title": "Unit Titles Act 2001", "slug": "unit-titles-act-2001",
     "citation": "A2001-16", "kind": "Act"},
    {"pdf": "2001-58.PDF", "title": "Community Title Act 2001", "slug": "community-title-act-2001",
     "citation": "A2001-58", "kind": "Act"},
    {"pdf": "2011-39.PDF", "title": "Unit Titles (Management) Regulation 2011", "slug": "unit-titles-management-regulation-2011",
     "citation": "SL2011-39", "kind": "Regulation"},
    {"pdf": "2011-41.PDF", "title": "Unit Titles (Management) Act 2011", "slug": "unit-titles-management-act-2011",
     "citation": "A2011-41", "kind": "Act"},
]

DEFINITIVE_RE = [
    re.compile(r'Authorised by the ACT Parliamentary Counsel'),
    re.compile(r'^Effective:\s'),
    re.compile(r'^R\d+$'),
    re.compile(r'^\d{1,2}/\d{1,2}/\d{2,4}$'),
    re.compile(r'^(page|contents)\s+\d+\b', re.I),
]

CANDIDATE_RE = re.compile(
    r'^(Dictionary|Endnotes)$'
    r'|Part\s+\d+[A-Za-z]*$'
    r'|Division\s+\d+(\.\d+)?[A-Za-z]*$'
    r'|Subdivision\s+\d+(\.\d+){1,2}[A-Za-z]*$'
    r'|Schedule\s+\d+[A-Za-z]*$'
    r'|^Section\s+\d+[A-Za-z]*$'
)

# Titles always start with a capital letter; requiring that rules out
# sentences that merely mention the Part/Division/Schedule mid-clause, e.g.
# "Schedule 3 applies to general meetings ..." (a cross-reference, not a
# heading — the giveaway is the lowercase verb straight after the number).
PART_RE = re.compile(r'^Part\s+(\d+[A-Za-z]*)\s+([A-Z]\S*.*)$')
# A schedule can contain its own "Part 1.1"-style sub-parts (e.g. two separate
# codes of conduct in one schedule), each restarting clause numbering at 1.
SCHEDULE_PART_RE = re.compile(r'^Part\s+(\d+\.\d+)\s+([A-Z]\S*.*)$')
DIVISION_RE = re.compile(r'^Division\s+(\d+(?:\.\d+)?[A-Za-z]*)\s+([A-Z]\S*.*)$')
SUBDIVISION_RE = re.compile(r'^Subdivision\s+(\d+(?:\.\d+){1,2}[A-Za-z]*)\s+([A-Z]\S*.*)$')
SCHEDULE_RE = re.compile(r'^Schedule\s+(\d+[A-Za-z]*)\s+([A-Z]\S*.*)$')
NUMBERED_SECTION_RE = re.compile(r'^(\d+[A-Z]{0,2}(?:\.\d+[A-Z]?)?)\s+([A-Z][^\n]{0,140})$')
CLAUSE_RE = re.compile(r'^\(([a-zA-Z0-9]+)\)\s')
NOTE_RE = re.compile(r'^(Note\b|Notes\b|Example\b|Examples\b)')
# Inside "Endnotes", only these are real subsection headings — everything else
# that looks like "<number> <Capitalised words>" there is actually a fragment
# from the legislation/amendment-history tables (e.g. "1 July 2015").
ENDNOTE_SUBSECTION_RE = re.compile(
    r'^\d+\s+(About the endnotes|Abbreviation key|Legislation history|'
    r'Amendment history|Earlier republications|'
    r'Expired transitional or validating provisions)$'
)

def is_definitive(s, title):
    if s == title:
        return True
    for r in DEFINITIVE_RE:
        if r.search(s):
            return True
    return False

def clean_lines(raw_text, title):
    lines = raw_text.split("\n")
    n = len(lines)
    is_true = [False]*n
    for i, l in enumerate(lines):
        s = l.strip()
        if s and is_definitive(s, title):
            is_true[i] = True

    def next_nonblank(i):
        k = i+1
        while k < n and lines[k].strip() == '':
            k += 1
        return k if k < n else None

    changed = True
    rounds = 0
    while changed and rounds < 20:
        changed = False
        rounds += 1
        for i in range(n):
            if is_true[i]:
                continue
            s = lines[i].strip()
            if not s or not CANDIDATE_RE.search(s):
                continue
            k = next_nonblank(i)
            if k is not None and is_true[k]:
                is_true[i] = True
                changed = True

    kept = [lines[i] for i in range(n) if not is_true[i]]
    return "\n".join(kept)

def find_nth_exact_line(text, marker, n_target):
    pos = 0
    count = 0
    while True:
        pos = text.find(marker, pos)
        if pos == -1:
            return -1
        line_start = text.rfind("\n", 0, pos) + 1
        line_end = text.find("\n", pos)
        line = text[line_start:line_end if line_end != -1 else len(text)].strip()
        if line == marker:
            count += 1
            if count == n_target:
                return line_start
        pos += len(marker)

def extract_metadata(raw):
    head = raw[:700]
    def grab(pat):
        m = re.search(pat, head)
        return m.group(1).strip() if m else None
    return {
        "republication_no": grab(r'Republication No\s+(\d+)'),
        "effective": grab(r'Effective:\s*([^\n]+)'),
        "last_amendment": grab(r'Last amendment made by\s+([^\n]+)'),
    }

TOC_START_RE = re.compile(
    r'^(Part\s+\d+[A-Za-z]*|Division\s+\d+(?:\.\d+)?[A-Za-z]*|Subdivision\s+\d+(?:\.\d+){1,2}[A-Za-z]*|Schedule\s+\d+[A-Za-z]*|Dictionary|Endnotes)\b(.*)$'
)
TOC_NUMBERED_RE = re.compile(r'^\d+[A-Z]{0,2}(?:\.\d+[A-Z]?)?\s')

def extract_toc_titles(cleaned, toc_start, toc_end):
    """The front-matter Contents listing gives the full, unwrapped title for
    each Part/Division/Schedule — the body text reprints these as running-page
    headers and sometimes line-wraps them, so the Contents copy is used as the
    source of truth for heading labels."""
    lines = [l.strip() for l in cleaned[toc_start:toc_end].split("\n")]
    n = len(lines)
    titles = {}

    def ends_with_pagenum(text):
        return re.search(r'\d+$', text) is not None

    i = 0
    while i < n:
        s = lines[i]
        if not s:
            i += 1
            continue
        m = TOC_START_RE.match(s)
        if m:
            kind_num = m.group(1)
            full = s
            j = i + 1
            while j < n and lines[j] and not TOC_START_RE.match(lines[j]) and not TOC_NUMBERED_RE.match(lines[j]):
                full += " " + lines[j]
                j += 1
            full_clean = re.sub(r'\s+\d+$', '', full).strip()
            if kind_num.startswith("Part"):
                hid = slugify_heading("part", kind_num.split(None, 1)[1])
            elif kind_num.startswith("Subdivision"):
                hid = slugify_heading("subdiv", kind_num.split(None, 1)[1])
            elif kind_num.startswith("Division"):
                hid = slugify_heading("div", kind_num.split(None, 1)[1])
            elif kind_num.startswith("Schedule"):
                hid = slugify_heading("sch", kind_num.split(None, 1)[1])
            else:
                hid = kind_num.lower()
            if hid not in titles:
                titles[hid] = full_clean
            i = j if j > i else i + 1
            continue
        m2 = TOC_NUMBERED_RE.match(s)
        if m2:
            num = s.split(None, 1)[0]
            full = s
            j = i + 1
            while not ends_with_pagenum(full) and j < n and lines[j] and not TOC_START_RE.match(lines[j]) and not TOC_NUMBERED_RE.match(lines[j]):
                full += " " + lines[j]
                j += 1
            full_clean = re.sub(r'\s+\d+$', '', full).strip()
            # drop the leading number token to get just the title, matching
            # the NUMBERED_SECTION_RE grouping used for body headings
            full_clean = re.sub(r'^\S+\s+', '', full_clean, count=1)
            hid = slugify_heading("s", num)
            if hid not in titles:
                titles[hid] = f"{num} {full_clean}"
            i = j if j > i else i + 1
            continue
        i += 1
    return titles

def slugify_heading(kind, num):
    num_s = re.sub(r'[.\s]+', '-', num)
    return f"{kind}-{num_s}"

def parse_body(body):
    lines = body.split("\n")
    blocks = []
    current = None  # dict(type, text list)
    in_endnotes = False
    # Schedules restart their own clause numbering at 1, which otherwise
    # collides with the Act's (unrelated) main section numbers of the same
    # value. Scope numbered-heading ids to the enclosing schedule to keep
    # anchors unique.
    schedule_scope = ""
    scope_prefix = ""

    def close():
        nonlocal current
        if current is not None and current["text"]:
            blocks.append(current)
        current = None

    for raw_line in lines:
        s = raw_line.strip()
        if s == "":
            close()
            continue

        m = PART_RE.match(s)
        if m:
            close()
            schedule_scope = ""
            scope_prefix = ""
            blocks.append({"type": "h2", "kind": "struct", "text": [s], "id": slugify_heading("part", m.group(1))})
            continue
        m = SCHEDULE_RE.match(s)
        if m:
            close()
            sch_id = slugify_heading("sch", m.group(1))
            schedule_scope = sch_id + "-"
            scope_prefix = schedule_scope
            blocks.append({"type": "h2", "kind": "struct", "text": [s], "id": sch_id})
            continue
        if s in ("Dictionary", "Endnotes"):
            close()
            schedule_scope = ""
            scope_prefix = ""
            blocks.append({"type": "h2", "kind": "struct", "text": [s], "id": s.lower()})
            in_endnotes = (s == "Endnotes")
            continue
        m = SCHEDULE_PART_RE.match(s)
        if m and schedule_scope:
            close()
            sub_id = schedule_scope + slugify_heading("part", m.group(1))
            scope_prefix = sub_id + "-"
            blocks.append({"type": "h3", "kind": "struct", "text": [s], "id": sub_id})
            continue
        m = DIVISION_RE.match(s)
        if m:
            close()
            blocks.append({"type": "h3", "kind": "struct", "text": [s], "id": slugify_heading("div", m.group(1))})
            continue
        m = SUBDIVISION_RE.match(s)
        if m:
            close()
            blocks.append({"type": "h4", "kind": "struct", "text": [s], "id": slugify_heading("subdiv", m.group(1))})
            continue
        if in_endnotes and ENDNOTE_SUBSECTION_RE.match(s):
            close()
            m = NUMBERED_SECTION_RE.match(s)
            # These repeat as a running-header reminder on every page of their
            # subsection (like Part/Division headings), so dedup like "struct".
            # Prefixed distinctly so they can't collide with a real numbered
            # section elsewhere in the Act that happens to share the number.
            blocks.append({"type": "h3", "kind": "struct", "text": [s], "id": slugify_heading("endnote", m.group(1))})
            continue
        m = NUMBERED_SECTION_RE.match(s)
        if m and not s.rstrip().endswith((",", ";", ":")) and not in_endnotes:
            close()
            blocks.append({"type": "h3", "kind": "section", "text": [s], "id": scope_prefix + slugify_heading("s", m.group(1))})
            continue
        m = CLAUSE_RE.match(s)
        if m:
            close()
            current = {"type": "clause", "text": [s]}
            continue
        if NOTE_RE.match(s):
            close()
            current = {"type": "note", "text": [s]}
            continue
        # continuation or new normal paragraph
        if current is None or current["type"] == "h2" or current["type"] == "h3":
            current = {"type": "para", "text": [s]}
        else:
            current["text"].append(s)
    close()

    # De-duplicate structural headings (Part/Division/Schedule/Dictionary/Endnotes):
    # the source PDF reprints the current Part/Division as a running-header
    # reminder on every page, so only the first occurrence of each is real.
    seen_struct_ids = set()
    deduped = []
    for b in blocks:
        if b.get("kind") == "struct":
            if b["id"] in seen_struct_ids:
                continue
            seen_struct_ids.add(b["id"])
        deduped.append(b)
    return deduped

def render_html(doc, blocks, meta):
    parts_nav = []
    body_html = []
    for b in blocks:
        text = " ".join(b["text"])
        esc = html.escape(text)
        if b["type"] == "h2":
            body_html.append(f'<h2 id="{b["id"]}">{esc}</h2>')
            parts_nav.append((b["id"], esc))
        elif b["type"] == "h3":
            body_html.append(f'<h3 id="{b["id"]}">{esc}</h3>')
        elif b["type"] == "h4":
            body_html.append(f'<h4 id="{b["id"]}">{esc}</h4>')
        elif b["type"] == "clause":
            body_html.append(f'<p class="clause">{esc}</p>')
        elif b["type"] == "note":
            body_html.append(f'<p class="note">{esc}</p>')
        else:
            body_html.append(f'<p>{esc}</p>')

    nav_items = "\n".join(f'<li><a href="#{i}">{t}</a></li>' for i, t in parts_nav)

    meta_bits = []
    if meta.get("effective"):
        meta_bits.append(f'Effective: {html.escape(meta["effective"])}')
    if meta.get("republication_no"):
        meta_bits.append(f'Republication No {html.escape(meta["republication_no"])}')
    if meta.get("last_amendment"):
        meta_bits.append(f'Last amendment: {html.escape(meta["last_amendment"])}')
    meta_line = " &middot; ".join(meta_bits)

    other_docs = "\n".join(
        f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])}</a></li>'
        for d in DOCS if d["slug"] != doc["slug"]
    )

    return f"""<!doctype html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<title>{html.escape(doc["title"])} — Strata Bot Knowledge Base</title>
<meta name="description" content="Full text of the {html.escape(doc["title"])} ({doc["citation"]}), Australian Capital Territory.">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="style.css">
</head>
<body>
<a class="skip-link" href="#main">Skip to content</a>
<header class="site-header">
<p class="site-name"><a href="index.html">Strata Bot Knowledge Base</a></p>
</header>
<nav class="doc-switch" aria-label="Other documents">
<ul>
{other_docs}
</ul>
</nav>
<main id="main">
<h1>{html.escape(doc["title"])}</h1>
<p class="doc-meta">{doc["citation"]} ({doc["kind"]}), Australian Capital Territory{" &middot; " + meta_line if meta_line else ""}</p>
<p class="source-note">This page reproduces the full text of the <a href="https://www.legislation.act.gov.au/">ACT legislation register</a> republication for reference by an automated assistant. It is not an authorised legal copy &mdash; always confirm current provisions at <a href="https://www.legislation.act.gov.au/">legislation.act.gov.au</a>.</p>
<nav class="toc" aria-label="Table of contents">
<h2 class="toc-heading">Contents</h2>
<ul>
{nav_items}
</ul>
</nav>
<article>
{chr(10).join(body_html)}
</article>
</main>
<footer class="site-footer">
<p><a href="index.html">&larr; Back to knowledge base index</a></p>
</footer>
</body>
</html>
"""

def build_index():
    items = []
    for d in DOCS:
        items.append(f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])}</a> <span class="citation">({d["citation"]})</span></li>')
    items_html = "\n".join(items)
    return f"""<!doctype html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<title>Strata Bot Knowledge Base</title>
<meta name="description" content="Reference knowledge base of ACT unit titles and community title legislation.">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="style.css">
</head>
<body>
<a class="skip-link" href="#main">Skip to content</a>
<header class="site-header">
<p class="site-name"><a href="index.html">Strata Bot Knowledge Base</a></p>
</header>
<main id="main">
<h1>Strata Bot Knowledge Base</h1>
<p>Reference text of the Australian Capital Territory legislation governing unit titles, community title, and owners corporation management. Each document below is reproduced in full on its own page.</p>
<ul class="doc-list">
{items_html}
</ul>
<p class="source-note">Source: <a href="https://www.legislation.act.gov.au/">ACT legislation register</a>. These pages are unofficial reproductions kept for use as a reference knowledge base; always confirm current provisions at the official register.</p>
</main>
</body>
</html>
"""

CSS = """:root {
  color-scheme: light dark;
  --bg: #ffffff;
  --fg: #1a1a1a;
  --muted: #555555;
  --link: #0b4f9c;
  --border: #d8d8d8;
  --accent-bg: #f4f6f8;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #14171a;
    --fg: #e8e8e8;
    --muted: #a8adb2;
    --link: #7fb2ff;
    --border: #33373b;
    --accent-bg: #1d2124;
  }
}
* { box-sizing: border-box; }
body {
  background: var(--bg);
  color: var(--fg);
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  line-height: 1.55;
  margin: 0;
  padding: 0;
}
.skip-link {
  position: absolute;
  left: -999px;
  top: 0;
  background: var(--accent-bg);
  color: var(--fg);
  padding: 0.5rem 1rem;
}
.skip-link:focus { left: 0; z-index: 10; }
.site-header {
  border-bottom: 1px solid var(--border);
  padding: 0.75rem 1.25rem;
}
.site-name a { font-weight: 700; text-decoration: none; color: var(--fg); }
main {
  max-width: 46rem;
  margin: 0 auto;
  padding: 1.5rem 1.25rem 4rem;
}
h1, h2, h3 { line-height: 1.25; }
h1 { font-size: 1.7rem; margin-top: 0.5rem; }
h2 { font-size: 1.3rem; margin-top: 2.2rem; border-top: 1px solid var(--border); padding-top: 1.2rem; }
h3 { font-size: 1.05rem; margin-top: 1.4rem; }
h4 { font-size: 0.98rem; margin-top: 1.2rem; color: var(--muted); }
a { color: var(--link); }
p { margin: 0.8rem 0; }
p.clause { margin-left: 1.5rem; }
p.note, p.source-note, p.doc-meta { color: var(--muted); font-size: 0.92rem; }
p.note { margin-left: 1.5rem; font-style: italic; }
.doc-switch {
  max-width: 46rem;
  margin: 0.75rem auto 0;
  padding: 0 1.25rem;
}
.doc-switch ul { display: flex; flex-wrap: wrap; gap: 0.75rem 1.25rem; list-style: none; padding: 0; margin: 0; font-size: 0.9rem; }
.toc {
  background: var(--accent-bg);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 1rem 1.25rem;
  margin: 1.5rem 0;
}
.toc-heading { margin: 0 0 0.5rem; border: none; padding: 0; font-size: 1rem; }
.toc ul { columns: 1; margin: 0; padding-left: 1.1rem; }
.toc li { break-inside: avoid; margin: 0.15rem 0; }
.doc-list { list-style: none; padding: 0; }
.doc-list li { padding: 0.6rem 0; border-bottom: 1px solid var(--border); }
.citation { color: var(--muted); font-size: 0.9rem; }
.site-footer { max-width: 46rem; margin: 0 auto; padding: 1rem 1.25rem 3rem; }
"""

def main():
    for doc in DOCS:
        path = f"{SRC_DIR}/{doc['pdf']}"
        raw = pdf_to_text(path)
        meta = extract_metadata(raw)
        cleaned = clean_lines(raw, doc["title"])
        toc_start = find_nth_exact_line(cleaned, "Australian Capital Territory", 2)
        start = find_nth_exact_line(cleaned, "Australian Capital Territory", 3)
        body = cleaned[start:] if start != -1 else cleaned
        end = body.find("\xa9  Australian Capital Territory")
        if end == -1:
            end = body.find("©  Australian Capital Territory")
        if end != -1:
            body = body[:end]
        toc_titles = extract_toc_titles(cleaned, toc_start, start) if toc_start != -1 and start != -1 else {}
        blocks = parse_body(body)
        seen_section_ids = set()
        for b in blocks:
            kind = b.get("kind")
            if kind == "section":
                # Numbers can coincidentally repeat (e.g. an inline numbered
                # example gets misread as a heading with the same id as a real
                # section). Only trust the TOC swap-in for the first, and only
                # when the TOC text is a superset of what's already there, so
                # a genuine mismatch is left alone rather than overwritten.
                first_seen = b["id"] not in seen_section_ids
                seen_section_ids.add(b["id"])
                if first_seen and b["id"] in toc_titles:
                    existing = b["text"][0]
                    candidate = toc_titles[b["id"]]
                    if candidate.startswith(existing) or existing.startswith(candidate):
                        b["text"] = [candidate if len(candidate) >= len(existing) else existing]
            elif kind == "struct" and b["id"] in toc_titles:
                b["text"] = [toc_titles[b["id"]]]
        html_out = render_html(doc, blocks, meta)
        with open(f"{OUT_DIR}/{doc['slug']}.html", "w") as f:
            f.write(html_out)
        h2_count = sum(1 for b in blocks if b["type"] == "h2")
        h3_count = sum(1 for b in blocks if b["type"] == "h3")
        print(doc["slug"], "blocks=", len(blocks), "h2=", h2_count, "h3=", h3_count, "meta=", meta)

    with open(f"{OUT_DIR}/index.html", "w") as f:
        f.write(build_index())
    with open(f"{OUT_DIR}/style.css", "w") as f:
        f.write(CSS)

if __name__ == "__main__":
    main()
