# Strata Bot Knowledge Base

A simple, static reference site reproducing the full text of the ACT
(Australian Capital Territory) legislation governing unit titles, community
title, and owners corporation management. Built to be used as a knowledge
base for an automated assistant (e.g. a Copilot agent) — one plain HTML page
per document, semantic headings, no JavaScript.

Live site: published via GitHub Pages from this repo's `main` branch.

## Documents

| Page | Citation | Source |
|---|---|---|
| `unit-titles-act-2001.html` | A2001-16 | [Unit Titles Act 2001](https://www.legislation.act.gov.au/a/2001-16/) |
| `community-title-act-2001.html` | A2001-58 | [Community Title Act 2001](https://www.legislation.act.gov.au/a/2001-58/) |
| `unit-titles-management-regulation-2011.html` | SL2011-39 | [Unit Titles (Management) Regulation 2011](https://www.legislation.act.gov.au/sl/2011-39/) |
| `unit-titles-management-act-2011.html` | A2011-41 | [Unit Titles (Management) Act 2011](https://www.legislation.act.gov.au/a/2011-41/) |

These are unofficial reproductions kept for reference. They are **not**
authorised legal copies — always confirm current provisions at the
[ACT legislation register](https://www.legislation.act.gov.au/).

## How it's built

- `source-docs/` holds the Word (`.docx`) republication of each document,
  downloaded from the ACT legislation register's "Download DOCX" link.
- `scripts/build.py` reads each `.docx` with `python-docx` and maps its
  named paragraph styles (`A H2 Part`, `A H5 Sec`, `Sch clause heading`,
  `AmdtsEntries`, etc. — the ACT Parliamentary Counsel's Word template)
  straight onto HTML headings and paragraph classes, then writes the HTML
  pages plus `index.html`/`style.css` to the repo root.
- Headings are emitted at their real depth — Part/Schedule at `h2`,
  Division/Schedule Part at `h3`, Subdivision at `h4`, and each section or
  schedule clause one level below whatever contains it — so the page outline
  matches the document's own hierarchy.
- Because schedule clause numbering restarts in each Part (the UTMA's
  schedule 1 has a clause 8 in both part 1.1 and part 1.2), and a bare
  "26 Other qualifications…" heading doesn't say which Part it sits under,
  every section and clause heading carries a small bracketed locator after
  the verbatim heading text. The legislative wording itself is untouched.
- Tables are walked in document order alongside the paragraphs. The
  reviewable-decisions schedules of the Unit Titles Act and Community Title
  Act exist only as tables, as do the endnote abbreviation key and the list
  of earlier republications.

The DOCX republication was used instead of the PDF (also available from the
register) because its paragraphs already carry the real document structure:
Part/Division/Schedule headings are distinct styles rather than text
reprinted as a running page header, no heading text is wrapped across a
page break, and Schedule content is tagged with its own styles so its
independent clause numbering can't collide with the main numbering. The
PDF's per-page header/footer boilerplate and line-wrapped headings would
otherwise need to be reconstructed heuristically.

To regenerate (e.g. after ACT publishes a newer republication — download the
new DOCX into `source-docs/` first, keeping the same filename):

```
pip install python-docx
python3 scripts/build.py
```

The script prints a warning listing any paragraph style it doesn't
recognise, so a template change in a future republication won't be silently
dropped.
