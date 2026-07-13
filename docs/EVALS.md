# Rendering parity eval

DocxInWeb renders `.docx` natively in the browser — no server round-trip, no
image conversion. The parity eval measures how close that render is to the
reference, which is **desktop Microsoft Word** itself.

## What it measures

Every fixture is a real `.docx`. For each one:

1. Desktop Microsoft Word exports the file to PDF — this is the ground truth.
2. DocxInWeb renders the same file in a headless browser.
3. Both the Word PDF and the DocxInWeb render are rasterized to PNGs at the same
   DPI and compared **page-by-page**.

The headline metric is **structural severity**: after registering a single
global page offset, it is the fraction of binary ink that still has no
counterpart within a small local tolerance, plus ink that only matches after a
one-line vertical shift beyond the calibrated noise floor. In other words, pure
antialiasing noise and sub-pixel line shifts do **not** count — only genuine
structural divergence (missing/extra ink, misplaced content) does. Each page is
also assigned a `driftClass` (clean / alignment / weight / colour / structural)
so alignment-, weight- and colour-only differences are named rather than folded
into the score.

## Categories

Pages are grouped into capability areas so strengths and weak spots are visible
at a glance rather than buried in a flat 1000-row list. The mapping lives in
`FIXTURE_CATEGORY` in `scripts/parity-report.mjs` (one line per fixture):

- **General word processing** — body text, notes, fields, TOC, tracked changes, forms
- **Complex tables** — nesting, row-splitting, computed heights, extreme/long tables
- **Math & equations** — OMML: inline/display equations, matrices, fractions
- **Other languages & scripts** — Arabic/kashida, Hebrew/bidi, Indic, Thai, CJK, ruby, vertical
- **Formatting & layout** — columns, sections, breaks, tabs, drop caps, borders, styles, frames
- **Graphics & media** — pictures, watermarks, WordArt, autofit shapes, emoji
- **Real-world documents** — the `wild-*` / `wild2-*` corpus of whole authored documents
- **Uncategorized** — any fixture not yet mapped (surfaces automatically so it gets classified)

Each category card shows page count, mean severity, worst page (linked to its
diff image) and how many pages exceed 1%.

## Provenance

The **Word** tab is the parity target: Word-authored files rendered against
Word's own PDF export. A separate **LibreOffice** tab tracks LibreOffice-authored
references as a distinct drift axis (currently deferred) so cross-suite
differences never pollute the Word headline.

## Known floors

Some residual severity is irreducible: Chrome and Word rasterize and antialias
glyphs differently, so even a pixel-perfect layout leaves a small
font-rasterization difference. Complex scripts (Arabic kashida, Indic shaping)
sit highest because their glyph shaping differs most between the two engines.
These floors are documented, not bugs to chase to zero.

## How to run it

```bash
# Full parallel run — writes parity/out/results.json (fast, all fixtures)
node scripts/parity-parallel.mjs

# Faster iteration: skip the slow appearance/semantic passes
DXW_PARITY_FAST=1 node scripts/parity-parallel.mjs

# Regenerate the HTML dashboard + report.png from existing results
# (parity-parallel writes results.json only; this rebuilds the report)
node scripts/parity-render-report.mjs
```

The dashboard is written to `parity/out/report.html` (and screenshotted to
`parity/out/report.png` at 1200px). The serial `scripts/parity-compare.mjs`
builds the same HTML inline at the end of a run; both callers share
`buildReport()` in `scripts/parity-report.mjs`, so the two paths produce
identical output.
