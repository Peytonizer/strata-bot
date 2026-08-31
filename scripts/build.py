import re, html, os
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if "__file__" in dir() else "."
SRC_DIR = os.path.join(REPO_ROOT, "source-docs")
OUT_DIR = REPO_ROOT

DOCS = [
    {"docx": "2001-16.docx", "title": "Unit Titles Act 2001", "slug": "unit-titles-act-2001",
     "citation": "A2001-16", "kind": "Act", "abbrev": "UTA"},
    {"docx": "2001-58.docx", "title": "Community Title Act 2001", "slug": "community-title-act-2001",
     "citation": "A2001-58", "kind": "Act", "abbrev": "CTA"},
    {"docx": "2011-39.docx", "title": "Unit Titles (Management) Regulation 2011", "slug": "unit-titles-management-regulation-2011",
     "citation": "SL2011-39", "kind": "Regulation", "abbrev": "UTMR"},
    {"docx": "2011-41.docx", "title": "Unit Titles (Management) Act 2011", "slug": "unit-titles-management-act-2011",
     "citation": "A2011-41", "kind": "Act", "abbrev": "UTMA"},
]

# ACT Parliamentary Counsel's Word template uses named paragraph styles that
# directly encode document structure — no running headers/footers, no
# line-wrapped headings, no sentence/heading ambiguity to guess at. This maps
# every style seen across all 4 documents to how it should be handled.

# style name -> (heading level, id-kind). Levels are the real document
# hierarchy, so a section always nests one level below the Part/Division/
# Subdivision (or Schedule/Schedule Part) that contains it.
STRUCT_STYLES = {
    "A H2 Part": (2, "part"),
    "A H3 Div": (3, "div"),
    "A H4 SubDiv": (4, "subdiv"),
    "Sched-heading": (2, "sch"),
    "Sched-Part": (3, "schpart"),
    "Dict-Heading": (2, "fixed:dictionary"),
    "Endnote1": (2, "fixed:endnotes"),
    "Endnote2": (3, "endnote"),
}

SECTION_STYLES = {"A H5 Sec", "Sch clause heading"}

CSS_CLASS_BY_STYLE = {
    "A main": "lead", "Sch A main": "lead", "LongTitle": "lead",
    "A main return": "lead lead-return",
    "A para": "clause clause-a", "Sch A para": "clause clause-a",
    "A subpara": "clause clause-i", "Sch A subpara": "clause clause-i",
    "A subsubpara": "clause clause-A",
    "aDef": "def", "aDef para": "def", "aDef subpara": "def",
    "ref": "ref",
    "Penalty": "penalty",
    "New Act": "amdt-act", "New Reg": "amdt-act",
    "Act details": "amdt-detail", "As am by": "amdt-detail",
    "AmdtsEntryHd": "amdt-head",
    "AmdtsEntries": "amdt-entry", "AmdtsEntriesDefL2": "amdt-entry",
    "EndNoteTextPub": "", "EndNoteTextEPS": "",
    "Normal": "",
    "Formula": "",
}
# Any style starting with these prefixes gets the given class.
CSS_CLASS_PREFIX = {
    "aNote": "note",
    "aExam": "example",
}

# Styles that never contribute body content (front matter, table of
# contents, empty structural bookmarks).
SKIP_STYLES = {
    "Billname", "Billname1", "BillBasic", "ActNo", "RepubNo", "EffectiveDate",
    "CoverInForce", "CoverHeading", "CoverSubHdg", "CoverText", "CoverTextBullet",
    "CoverActName", "N-TOCheading", "N-9pt", "PageBreak", "Placeholder",
    "00SigningPage", "01Contents", "02Text", "03Schedule", "04Dictionary",
    "05EndNote", "06Copyright", "N-line3",
}

def is_skippable(style_name, text):
    if style_name.startswith("toc"):
        return True
    if style_name in SKIP_STYLES:
        return True
    if style_name == "Normal" and (text == "Australian Capital Territory" or text.startswith("©")):
        return True
    return False

def slugify_num(num):
    return re.sub(r'[.\s]+', '-', num)

def split_num_title(text):
    """Heading paragraphs are stored as '<number>\\t<title>'."""
    if "\t" in text:
        num, title = text.split("\t", 1)
        return num.strip(), title.strip()
    parts = text.split(None, 1)
    if len(parts) == 2:
        return parts[0], parts[1]
    return text, ""

NUM_TAIL_RE = re.compile(r'(\d+(?:\.\d+)*[A-Za-z]*)$')

def bare_num(label):
    """Part/Division/Subdivision/Schedule headings store the label word in
    front of the number ('Division 3.2') — strip it for id-building."""
    m = NUM_TAIL_RE.search(label)
    return m.group(1) if m else label

def extract_metadata(doc):
    meta = {}
    for p in doc.paragraphs:
        t = p.text.strip()
        if p.style.name == "RepubNo":
            m = re.search(r'\d+', t)
            if m:
                meta["republication_no"] = m.group(0)
        elif p.style.name == "EffectiveDate":
            meta["effective"] = t.split(":", 1)[-1].strip().replace("\xa0", " ")
        elif p.style.name == "CoverInForce" and t.lower().startswith("last amendment made by"):
            meta["last_amendment"] = t.split("by", 1)[-1].strip()
    return meta

def iter_body(doc):
    """Paragraphs and tables in document order. Schedules of reviewable
    decisions, the endnote abbreviation key and the earlier-republications
    list are all tables, so walking doc.paragraphs alone drops them."""
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)

def is_header_style(style_name):
    return style_name.endswith("ColHd") or style_name.endswith("Hdg")

def read_table(table):
    """Rows of cells, each cell a list of paragraph strings, plus a flag for
    whether the row is a column-heading row."""
    rows = []
    for row in table.rows:
        cells, seen = [], set()
        for cell in row.cells:
            key = id(cell._tc)
            if key in seen:      # horizontally merged cell, already emitted
                continue
            seen.add(key)
            lines = []
            for p in cell.paragraphs:
                t = p.text.strip().replace("\t", " ")
                if t:
                    lines.append(t)
            cells.append(lines)
        if not any(cells):
            continue
        styles = [p.style.name for c in row.cells for p in c.paragraphs if p.text.strip()]
        rows.append({"cells": cells, "header": bool(styles) and all(is_header_style(s) for s in styles)})
    return rows

def parse_doc(doc, unmapped_styles):
    blocks = []
    schedule_scope = ""   # e.g. "sch-1-"
    subpart_scope = ""    # e.g. "part-1-1-" (nested inside a schedule)
    scope = []            # [(level, "Part 3"), (level, "Division 3.3")]
    seen_struct_ids = {}

    def unique_id(hid):
        n = seen_struct_ids.get(hid, 0) + 1
        seen_struct_ids[hid] = n
        return hid if n == 1 else f"{hid}-{n}"

    for item in iter_body(doc):
        if isinstance(item, Table):
            rows = read_table(item)
            if rows:
                blocks.append({"type": "table", "kind": "table", "rows": rows})
            continue

        p = item
        text = p.text.strip()
        if not text:
            continue
        style = p.style.name
        if is_skippable(style, text):
            continue

        if style in STRUCT_STYLES:
            level, kind = STRUCT_STYLES[style]
            label = None
            if kind == "part":
                schedule_scope = subpart_scope = ""
                num, title = split_num_title(text)
                hid = f"part-{slugify_num(bare_num(num))}"
                label = num
            elif kind == "div":
                num, title = split_num_title(text)
                hid = f"div-{slugify_num(bare_num(num))}"
                label = num
            elif kind == "subdiv":
                num, title = split_num_title(text)
                hid = f"subdiv-{slugify_num(bare_num(num))}"
                label = num
            elif kind == "sch":
                num, title = split_num_title(text)
                subpart_scope = ""
                hid = f"sch-{slugify_num(bare_num(num))}"
                schedule_scope = hid + "-"
                label = num
            elif kind == "schpart":
                num, title = split_num_title(text)
                part_num = slugify_num(bare_num(num))
                hid = schedule_scope + f"part-{part_num}"
                subpart_scope = f"part-{part_num}-"
                label = num
            elif kind == "fixed:dictionary":
                title, hid = "Dictionary", "dictionary"
                schedule_scope = subpart_scope = ""
            elif kind == "fixed:endnotes":
                title, hid = "Endnotes", "endnotes"
                schedule_scope = subpart_scope = ""
            elif kind == "endnote":
                num, t = split_num_title(text)
                title = f"{num} {t}"
                hid = f"endnote-{slugify_num(num)}"

            while scope and scope[-1][0] >= level:
                scope.pop()
            if label:
                scope.append((level, label))

            display = text.replace("\t", " ") if kind not in ("fixed:dictionary", "fixed:endnotes") else title
            blocks.append({"type": f"h{level}", "kind": "struct", "text": display,
                           "id": unique_id(hid), "level": level})
            continue

        if style in SECTION_STYLES:
            num, title = split_num_title(text)
            hid = schedule_scope + subpart_scope + f"s-{slugify_num(num)}"
            # Schedule clause numbering restarts in each Part (UTMA sch 1 has a
            # clause 8 in both pt 1.1 and pt 1.2), and a bare "26 ..." heading
            # doesn't say which Part or Division it sits under. The breadcrumb
            # keeps every heading self-identifying when it is read out of context.
            level = min(5, (scope[-1][0] + 1) if scope else 3)
            blocks.append({"type": f"h{level}", "kind": "section", "text": f"{num} {title}",
                           "id": unique_id(hid), "level": level,
                           "scope": ", ".join(lbl for _, lbl in scope)})
            continue

        cls = CSS_CLASS_BY_STYLE.get(style)
        if cls is None:
            for prefix, c in CSS_CLASS_PREFIX.items():
                if style.startswith(prefix):
                    cls = c
                    break
        if cls is None:
            unmapped_styles.add(style)
            cls = ""
        blocks.append({"type": "p", "kind": "content", "text": text.replace("\t", " "), "class": cls})

    return blocks

def render_toc(nav_items):
    """Nested list of the structural skeleton (Parts, Divisions, Subdivisions,
    Schedules, Schedule Parts, Dictionary, Endnotes). Sections are deliberately
    left out: repeating 250-odd section titles here would compete with the
    provisions themselves when the page is indexed."""
    out, depth, open_li = ["<ul>"], 1, False
    for hid, text, level in nav_items:
        while depth < level - 1:
            out.append("<ul>")       # nested list belongs inside the open <li>
            depth += 1
            open_li = False
        while depth > level - 1:
            if open_li:
                out.append("</li>")
                open_li = False
            out.append("</ul>")
            depth -= 1
            open_li = True           # the <li> that wrapped the nested list
        if open_li:
            out.append("</li>")
        out.append(f'<li><a href="#{hid}">{text}</a>')
        open_li = True
    while depth > 0:
        if open_li:
            out.append("</li>")
            open_li = False
        out.append("</ul>")
        depth -= 1
        open_li = depth > 0
    return "\n".join(out)

def render_table(rows):
    out = ['<div class="table-wrap">', "<table>"]
    body_open = False
    for i, row in enumerate(rows):
        if row["header"]:
            if i == 0:
                out.append("<thead>")
            tag = "th"
        else:
            if i > 0 and rows[0]["header"] and not body_open:
                out.append("</thead>")
                out.append("<tbody>")
                body_open = True
            tag = "td"
        cells = "".join(
            f"<{tag}>" + "<br>".join(html.escape(line) for line in cell) + f"</{tag}>"
            for cell in row["cells"]
        )
        out.append(f"<tr>{cells}</tr>")
    if body_open:
        out.append("</tbody>")
    out.append("</table>")
    out.append("</div>")
    return "\n".join(out)

def render_html(doc_meta, blocks, page_meta):
    nav_items = []
    body_html = []
    for b in blocks:
        if b["kind"] == "table":
            body_html.append(render_table(b["rows"]))
            continue
        esc = html.escape(b["text"])
        if b["kind"] in ("struct", "section"):
            tag = b["type"]
            scope = b.get("scope")
            extra = f' <span class="heading-scope">({html.escape(scope)})</span>' if scope else ""
            cls = ' class="struct"' if b["kind"] == "struct" else ""
            body_html.append(f'<{tag}{cls} id="{b["id"]}">{esc}{extra}</{tag}>')
            if b["kind"] == "struct":
                nav_items.append((b["id"], esc, b["level"]))
        else:
            cls = f' class="{b["class"]}"' if b["class"] else ""
            body_html.append(f'<p{cls}>{esc}</p>')

    nav_html = render_toc(nav_items)

    meta_bits = [f'{doc_meta["citation"]} ({doc_meta["kind"]})',
                 f'cited here as {doc_meta["abbrev"]}',
                 "Australian Capital Territory"]
    if page_meta.get("effective"):
        meta_bits.append(f'Effective: {html.escape(page_meta["effective"])}')
    if page_meta.get("republication_no"):
        meta_bits.append(f'Republication No {html.escape(page_meta["republication_no"])}')
    if page_meta.get("last_amendment"):
        meta_bits.append(f'Last amendment: {html.escape(page_meta["last_amendment"])}')
    meta_line = " &middot; ".join(meta_bits)

    other_docs = "\n".join(
        f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])} ({d["abbrev"]})</a></li>'
        for d in DOCS if d["slug"] != doc_meta["slug"]
    )

    return f"""<!doctype html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<title>{html.escape(doc_meta["title"])} ({doc_meta["abbrev"]}) — Strata Bot Knowledge Base</title>
<meta name="description" content="Full text of the {html.escape(doc_meta["title"])} ({doc_meta["citation"]}), Australian Capital Territory, cited as {doc_meta["abbrev"]}.">
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
<h1>{html.escape(doc_meta["title"])} ({doc_meta["abbrev"]})</h1>
<p class="doc-meta">{meta_line}</p>
<p class="source-note">This page reproduces the full text of the <a href="https://www.legislation.act.gov.au/">ACT legislation register</a> republication for reference by an automated assistant. It is not an authorised legal copy &mdash; always confirm current provisions at <a href="https://www.legislation.act.gov.au/">legislation.act.gov.au</a>.</p>
<nav class="toc" aria-label="Table of contents">
<h2 class="toc-heading">Contents</h2>
{nav_html}
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
    items = "\n".join(
        f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])}</a> '
        f'<span class="citation">({d["citation"]}, cited as {d["abbrev"]})</span></li>'
        for d in DOCS
    )
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
{items}
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
  --struct: #0b4f9c;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #14171a;
    --fg: #e8e8e8;
    --muted: #a8adb2;
    --link: #7fb2ff;
    --border: #33373b;
    --accent-bg: #1d2124;
    --struct: #9ec3ff;
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
h1, h2, h3, h4, h5 { line-height: 1.25; }
h1 { font-size: 1.7rem; margin-top: 0.5rem; }
h2 { font-size: 1.35rem; margin-top: 2.4rem; border-top: 1px solid var(--border); padding-top: 1.2rem; }
h3 { font-size: 1.15rem; margin-top: 1.8rem; }
h4 { font-size: 1.05rem; margin-top: 1.5rem; }
h5 { font-size: 1rem; margin-top: 1.3rem; }
/* Part/Division/Subdivision/Schedule headings are dividers, not provisions —
   set them apart so a section never reads as a sibling of its own container. */
h3.struct, h4.struct, h5.struct {
  color: var(--struct);
  text-transform: uppercase;
  letter-spacing: 0.04em;
  font-size: 0.85rem;
  margin-top: 2rem;
  padding-bottom: 0.3rem;
  border-bottom: 1px solid var(--border);
}
.heading-scope { font-weight: 400; font-size: 0.8rem; color: var(--muted); }
a { color: var(--link); }
p { margin: 0.8rem 0; }
p.lead { margin-top: 1rem; }
p.lead-return { margin-top: 0.4rem; }
p.clause-a { margin-left: 1.5rem; }
p.clause-i { margin-left: 3rem; }
p.clause-A { margin-left: 4.5rem; }
p.def { margin-left: 1.5rem; font-style: italic; }
p.note, p.example { margin-left: 1.5rem; color: var(--muted); font-size: 0.92rem; font-style: italic; }
p.penalty { margin-left: 1.5rem; font-size: 0.95rem; }
p.ref { color: var(--muted); font-size: 0.85rem; margin: 0.2rem 0 0.8rem; }
p.amdt-act { font-weight: 600; margin-top: 1rem; }
p.amdt-head { font-weight: 600; margin-top: 0.8rem; }
p.amdt-detail, p.amdt-entry { color: var(--muted); font-size: 0.88rem; margin: 0.15rem 0 0.15rem 1.5rem; }
p.source-note, p.doc-meta { color: var(--muted); font-size: 0.92rem; }
.table-wrap { overflow-x: auto; margin: 1rem 0; }
table { border-collapse: collapse; width: 100%; font-size: 0.92rem; }
th, td { border: 1px solid var(--border); padding: 0.4rem 0.6rem; text-align: left; vertical-align: top; }
th { background: var(--accent-bg); font-weight: 600; }
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
.toc ul { margin: 0; padding-left: 1.1rem; }
.toc ul ul { padding-left: 1.2rem; font-size: 0.94rem; }
.toc li { margin: 0.15rem 0; }
.doc-list { list-style: none; padding: 0; }
.doc-list li { padding: 0.6rem 0; border-bottom: 1px solid var(--border); }
.citation { color: var(--muted); font-size: 0.9rem; }
.site-footer { max-width: 46rem; margin: 0 auto; padding: 1rem 1.25rem 3rem; }
"""

def main():
    unmapped_styles = set()
    for doc_meta in DOCS:
        path = os.path.join(SRC_DIR, doc_meta["docx"])
        doc = docx.Document(path)
        page_meta = extract_metadata(doc)
        blocks = parse_doc(doc, unmapped_styles)
        out_html = render_html(doc_meta, blocks, page_meta)
        with open(os.path.join(OUT_DIR, f"{doc_meta['slug']}.html"), "w") as f:
            f.write(out_html)
        counts = {t: sum(1 for b in blocks if b["type"] == t) for t in ("h2", "h3", "h4", "h5", "table")}
        print(doc_meta["slug"], "blocks=", len(blocks), counts, "meta=", page_meta)

    with open(os.path.join(OUT_DIR, "index.html"), "w") as f:
        f.write(build_index())
    with open(os.path.join(OUT_DIR, "style.css"), "w") as f:
        f.write(CSS)

    if unmapped_styles:
        print("\nWARNING: unmapped styles rendered as plain <p> (review CSS_CLASS_BY_STYLE):")
        for s in sorted(unmapped_styles):
            print(" -", s)
    else:
        print("\nAll styles mapped.")

if __name__ == "__main__":
    main()
