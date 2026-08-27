import re, html, os
import docx

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if "__file__" in dir() else "."
SRC_DIR = os.path.join(REPO_ROOT, "source-docs")
OUT_DIR = REPO_ROOT

DOCS = [
    {"docx": "2001-16.docx", "title": "Unit Titles Act 2001", "slug": "unit-titles-act-2001",
     "citation": "A2001-16", "kind": "Act"},
    {"docx": "2001-58.docx", "title": "Community Title Act 2001", "slug": "community-title-act-2001",
     "citation": "A2001-58", "kind": "Act"},
    {"docx": "2011-39.docx", "title": "Unit Titles (Management) Regulation 2011", "slug": "unit-titles-management-regulation-2011",
     "citation": "SL2011-39", "kind": "Regulation"},
    {"docx": "2011-41.docx", "title": "Unit Titles (Management) Act 2011", "slug": "unit-titles-management-act-2011",
     "citation": "A2011-41", "kind": "Act"},
]

# ACT Parliamentary Counsel's Word template uses named paragraph styles that
# directly encode document structure — no running headers/footers, no
# line-wrapped headings, no sentence/heading ambiguity to guess at. This maps
# every style seen across all 4 documents to how it should be handled.

STRUCT_STYLES = {
    # style name -> (html tag, id-kind)
    "A H2 Part": ("h2", "part"),
    "A H3 Div": ("h3", "div"),
    "A H4 SubDiv": ("h4", "subdiv"),
    "Sched-heading": ("h2", "sch"),
    "Sched-Part": ("h3", "schpart"),
    "Dict-Heading": ("h2", "fixed:dictionary"),
    "Endnote1": ("h2", "fixed:endnotes"),
    "Endnote2": ("h3", "endnote"),
}

SECTION_STYLES = {"A H5 Sec", "Sch clause heading"}

CSS_CLASS_BY_STYLE = {
    "A main": "lead", "A main return": "lead", "Sch A main": "lead", "LongTitle": "lead",
    "A para": "clause clause-a", "Sch A para": "clause clause-a",
    "A subpara": "clause clause-i", "Sch A subpara": "clause clause-i",
    "A subsubpara": "clause clause-A",
    "aDef": "def", "aDef para": "def", "aDef subpara": "def",
    "ref": "ref",
    "Penalty": "clause clause-a",
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
    "aExam": "note",
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

def parse_doc(doc, unmapped_styles):
    blocks = []
    schedule_scope = ""   # e.g. "sch-1-"
    subpart_scope = ""    # e.g. "part-1-1-" (nested inside a schedule)
    seen_struct_ids = set()

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = p.style.name
        if is_skippable(style, text):
            continue

        if style in STRUCT_STYLES:
            tag, kind = STRUCT_STYLES[style]
            if kind == "part":
                schedule_scope = ""
                subpart_scope = ""
                num, title = split_num_title(text)
                hid = f"part-{slugify_num(bare_num(num))}"
            elif kind == "div":
                num, title = split_num_title(text)
                hid = f"div-{slugify_num(bare_num(num))}"
            elif kind == "subdiv":
                num, title = split_num_title(text)
                hid = f"subdiv-{slugify_num(bare_num(num))}"
            elif kind == "sch":
                num, title = split_num_title(text)
                subpart_scope = ""
                hid = f"sch-{slugify_num(bare_num(num))}"
                schedule_scope = hid + "-"
            elif kind == "schpart":
                num, title = split_num_title(text)
                part_num = slugify_num(bare_num(num))
                hid = schedule_scope + f"part-{part_num}"
                subpart_scope = f"part-{part_num}-"
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
            if hid in seen_struct_ids:
                continue  # defensive; DOCX shouldn't repeat these
            seen_struct_ids.add(hid)
            display = text.replace("\t", " ") if kind not in ("fixed:dictionary", "fixed:endnotes") else title
            blocks.append({"type": tag, "kind": "struct", "text": display, "id": hid})
            continue

        if style in SECTION_STYLES:
            num, title = split_num_title(text)
            hid = schedule_scope + subpart_scope + f"s-{slugify_num(num)}"
            blocks.append({"type": "h3", "kind": "section", "text": f"{num} {title}", "id": hid})
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

def render_html(doc_meta, blocks, page_meta):
    nav_items = []
    body_html = []
    for b in blocks:
        esc = html.escape(b["text"])
        if b["kind"] in ("struct", "section"):
            tag = b["type"]
            body_html.append(f'<{tag} id="{b["id"]}">{esc}</{tag}>')
            if b["kind"] == "struct" and tag == "h2":
                nav_items.append((b["id"], esc))
        else:
            cls = f' class="{b["class"]}"' if b["class"] else ""
            body_html.append(f'<p{cls}>{esc}</p>')

    nav_html = "\n".join(f'<li><a href="#{i}">{t}</a></li>' for i, t in nav_items)

    meta_bits = []
    if page_meta.get("effective"):
        meta_bits.append(f'Effective: {html.escape(page_meta["effective"])}')
    if page_meta.get("republication_no"):
        meta_bits.append(f'Republication No {html.escape(page_meta["republication_no"])}')
    if page_meta.get("last_amendment"):
        meta_bits.append(f'Last amendment: {html.escape(page_meta["last_amendment"])}')
    meta_line = " &middot; ".join(meta_bits)

    other_docs = "\n".join(
        f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])}</a></li>'
        for d in DOCS if d["slug"] != doc_meta["slug"]
    )

    return f"""<!doctype html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<title>{html.escape(doc_meta["title"])} — Strata Bot Knowledge Base</title>
<meta name="description" content="Full text of the {html.escape(doc_meta["title"])} ({doc_meta["citation"]}), Australian Capital Territory.">
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
<h1>{html.escape(doc_meta["title"])}</h1>
<p class="doc-meta">{doc_meta["citation"]} ({doc_meta["kind"]}), Australian Capital Territory{" &middot; " + meta_line if meta_line else ""}</p>
<p class="source-note">This page reproduces the full text of the <a href="https://www.legislation.act.gov.au/">ACT legislation register</a> republication for reference by an automated assistant. It is not an authorised legal copy &mdash; always confirm current provisions at <a href="https://www.legislation.act.gov.au/">legislation.act.gov.au</a>.</p>
<nav class="toc" aria-label="Table of contents">
<h2 class="toc-heading">Contents</h2>
<ul>
{nav_html}
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
    items = "\n".join(
        f'<li><a href="{d["slug"]}.html">{html.escape(d["title"])}</a> <span class="citation">({d["citation"]})</span></li>'
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
h1, h2, h3, h4 { line-height: 1.25; }
h1 { font-size: 1.7rem; margin-top: 0.5rem; }
h2 { font-size: 1.3rem; margin-top: 2.2rem; border-top: 1px solid var(--border); padding-top: 1.2rem; }
h3 { font-size: 1.05rem; margin-top: 1.4rem; }
h4 { font-size: 0.98rem; margin-top: 1.2rem; color: var(--muted); }
a { color: var(--link); }
p { margin: 0.8rem 0; }
p.lead { margin-top: 1rem; }
p.clause-a { margin-left: 1.5rem; }
p.clause-i { margin-left: 3rem; }
p.clause-A { margin-left: 4.5rem; }
p.def { margin-left: 1.5rem; font-style: italic; }
p.note { margin-left: 1.5rem; color: var(--muted); font-size: 0.92rem; font-style: italic; }
p.ref { color: var(--muted); font-size: 0.85rem; margin: 0.2rem 0 0.8rem; }
p.amdt-act { font-weight: 600; margin-top: 1rem; }
p.amdt-head { font-weight: 600; margin-top: 0.8rem; }
p.amdt-detail, p.amdt-entry { color: var(--muted); font-size: 0.88rem; margin: 0.15rem 0 0.15rem 1.5rem; }
p.source-note, p.doc-meta { color: var(--muted); font-size: 0.92rem; }
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
        h2 = sum(1 for b in blocks if b["type"] == "h2")
        h3 = sum(1 for b in blocks if b["type"] == "h3")
        h4 = sum(1 for b in blocks if b["type"] == "h4")
        print(doc_meta["slug"], "blocks=", len(blocks), "h2=", h2, "h3=", h3, "h4=", h4, "meta=", page_meta)

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

