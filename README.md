# strata.noradz.io

The landing page for a small family of strata things, and the ACT strata legislation reproduced
in full underneath it.

Live at **[strata.noradz.io](https://strata.noradz.io)**. A static site, no JavaScript beyond a
light/dark toggle on the hub, published by GitHub Pages from `main`.

## Structure

```
index.html               the hub — generated from the strata-kit manifest
legislation/             the reader: an index and one page per instrument
assets/hub.css           hub styling (fonts + kit theme + hub rules, concatenated)
assets/legislation.css   reader styling, built the same way
assets/fonts/            self-hosted Fraunces and DM Sans, both OFL
assets/theme-toggle.js   copied from the kit at build time
<slug>.html              redirect stubs at the documents' former root-level URLs
scripts/build.py         builds all of the above
source-docs/             the .docx republications the reader is generated from
vendor/strata-kit/       submodule: the shared palette and the family navigation bar
```

Everything except the source documents, the fonts and this README is generated. Edit
`scripts/build.py`, not the HTML.

## The documents

| Page | Citation | Source |
|---|---|---|
| `legislation/unit-titles-act-2001.html` | A2001-16 | [Unit Titles Act 2001](https://www.legislation.act.gov.au/a/2001-16/) |
| `legislation/community-title-act-2001.html` | A2001-58 | [Community Title Act 2001](https://www.legislation.act.gov.au/a/2001-58/) |
| `legislation/unit-titles-management-regulation-2011.html` | SL2011-39 | [Unit Titles (Management) Regulation 2011](https://www.legislation.act.gov.au/sl/2011-39/) |
| `legislation/unit-titles-management-act-2011.html` | A2011-41 | [Unit Titles (Management) Act 2011](https://www.legislation.act.gov.au/a/2011-41/) |

These are unofficial reproductions kept for reference. They are **not** authorised legal copies
— always confirm current provisions at the
[ACT legislation register](https://www.legislation.act.gov.au/).

## Building

```sh
git clone --recurse-submodules https://github.com/Peytonizer/strata-bot.git
pip install -r requirements.txt
python3 scripts/build.py
```

An existing clone that predates the submodule needs `git submodule update --init` once.

- `source-docs/` holds the Word (`.docx`) republication of each document, downloaded from the
  ACT legislation register's "Download DOCX" link.
- `scripts/build.py` reads each `.docx` with `python-docx` and maps its named paragraph styles
  (`A H2 Part`, `A H5 Sec`, `Sch clause heading`, `AmdtsEntries`, etc. — the ACT Parliamentary
  Counsel's Word template) straight onto HTML headings and paragraph classes, then writes the
  reader pages, the hub, the stylesheets, the redirect stubs and the sitemap.
- Headings are emitted at their real depth — Part/Schedule at `h2`, Division/Schedule Part at
  `h3`, Subdivision at `h4`, and each section or schedule clause one level below whatever
  contains it — so the page outline matches the document's own hierarchy.
- Because schedule clause numbering restarts in each Part (the UTMA's schedule 1 has a clause 8
  in both part 1.1 and part 1.2), and a bare "26 Other qualifications…" heading doesn't say
  which Part it sits under, every section and clause heading carries a small bracketed locator.

## The hub

`index.html` is generated from `vendor/strata-kit/projects.json`, which is also what fills the
navigation bar at the top of every page here, on
[lodger](https://lodger.noradz.io) and on [former](https://former.noradz.io). Adding a project
to the family means adding an entry in
[strata-kit](https://github.com/Peytonizer/strata-kit), then in this repo:

```sh
git submodule update --remote vendor/strata-kit
python3 scripts/build.py
```

and committing both the moved pin and the rebuilt pages.

## The old URLs

The documents used to sit at the root — `/unit-titles-act-2001.html` and friends — before the
hub took that spot. Each of those paths is now a stub that redirects into `legislation/`,
carrying the fragment across, so a link to a specific section still lands on that section.
