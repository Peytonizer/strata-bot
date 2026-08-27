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

- `source-pdfs/` holds the original republished PDFs downloaded from the ACT
  legislation register.
- `scripts/build.py` extracts text from each PDF (via `pypdf`), strips the
  repeated running header/footer boilerplate every page carries, reconstructs
  headings (Part/Division/Subdivision/Schedule/numbered sections) and
  paragraphs, and writes the HTML pages plus `index.html`/`style.css` to the
  repo root.

To regenerate (e.g. after ACT publishes a newer republication — swap the PDF
in `source-pdfs/` first):

```
pip install pypdf
python3 scripts/build.py
```

The heading structure is derived heuristically from the PDF's own table of
contents and text layout; the underlying legislative text itself is a direct
extraction, but a small number of inline numbered examples may occasionally
be mistagged as headings rather than body text.
