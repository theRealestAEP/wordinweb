# DocxInWeb

Open-source, high-fidelity **Word/.docx viewer (and, soon, editor) for the web** — think Google Docs' rendering quality, embeddable as a single React component.

```tsx
import { DocxView } from "@docxinweb/react";

<DocxView source="/report.docx" zoom={1} />              // render-only viewer
<DocxView source="/report.docx" editable onReady={api => {
  // api.applyFormat({ bold: true }) on the current selection
  // api.save() → edited .docx bytes
}} />
```

Editing is strictly opt-in via the `editable` flag — the default is a pure
render-only viewer.

## Why another docx renderer?

Existing options either convert to lossy HTML flow (docx-preview, mammoth) or require a server. DocxInWeb takes the same approach Word itself does: it runs a **real layout engine** in the browser —

1. **Parse** the OOXML package into a typed document model (styles with `basedOn` chains, numbering, sections, headers/footers, themes, relationships).
2. **Lay out** with measured text: canvas-metric line breaking, line-spacing rules (`auto`/`atLeast`/`exact`), spacing before/after, widow/orphan control, real page boxes per `sectPr`.
3. **Render** absolutely positioned pages — the browser does zero reflow, so what the engine computed is exactly what you see.

That is what makes the fidelity-critical features work:

- **Page numbers computed properly** — `PAGE` / `NUMPAGES` fields are resolved during pagination, per page, including `pgNumType` start/format overrides.
- **Header/footer spacing done right** — headers are measured and placed at `headerDistance`; the body top is `max(marginTop, headerBottom)` exactly like Word, including first-page and even/odd variants.
- **Divider lines render properly** — paragraph borders (including the classic bottom-border horizontal rule) are drawn as positioned edges.
- **Accurate fonts** — theme font resolution (`majorFont`/`minorFont`), canvas `fontBoundingBox` metrics for Word-parity line heights.
- **Tables** — grid/pct widths, cell spans, borders with table/cell conflict resolution, header-row repetition across pages, shading, vertical alignment.
- **Lists** — full `numbering.xml` support: multi-level counters, restarts, roman/letter/decimal formats, bullet glyph mapping.
- **Columns, section breaks, justified text, tabs, images, hyperlinks.**

## Packages

| Package | What it is |
| --- | --- |
| `@docxinweb/core` | Parser + layout engine + DOM renderer. Framework-agnostic, zero deps besides `fflate`. |
| `@docxinweb/react` | `<DocxView />` React component. |
| `apps/demo` | Vite demo app with file-open and zoom. |

## Development

```bash
npm install
npm test                 # core unit tests (parser + layout, deterministic measurer)
npm run build            # build core + react
npm -w demo run fixtures # generate sample .docx fixtures
npm run dev              # demo app at http://localhost:5173
```

## Architecture

```
.docx bytes
  └─ Package (fflate unzip, OPC part resolution)
      └─ DocxDocument           parse/: document, styles, numbering, theme, sections, rels
          └─ layoutDocument()   layout/: inline breaking → pagination → header/footer passes
              └─ LayoutResult   absolutely positioned PageItems (text/rect/edge/image)
                  └─ renderToDom() / <DocxView/>
```

The layout output (`LaidOutPage[]` of `PageItem`s) is renderer-agnostic — a canvas or SVG/PDF backend can consume the same geometry, and the future editor will map DOM positions back to model offsets through it.

## Editing architecture

The OOXML tree is the **source of truth**. Model nodes keep `src` references
to their `w:r`/`w:t` elements; every rendered text span maps back to
`(run, w:t, char offset)`. Commands split runs at selection boundaries,
mutate `w:rPr`, re-parse (`doc.refresh()`), and re-lay out — full relayout is
single-digit milliseconds. `doc.save()` re-serializes only the XML parts we
model; **everything untouched round-trips byte-for-byte** (comments,
footnotes, custom XML, embedded fonts survive edits unscathed).

Supported today: bold, italic, underline, strike, font size, font family,
text color, highlight — over any selection, including partial runs.

## Rendering parity

Fidelity is measured against **desktop Microsoft Word** itself. Every fixture is
a real `.docx`: Word exports it to PDF (the ground truth), DocxInWeb renders the
same file in the browser, and both are rasterized and compared **page-by-page**.
The headline metric is *structural severity* — the fraction of ink with no
counterpart after a single global page alignment, so antialiasing noise and
sub-pixel line shifts don't count. Pages are grouped into capability areas
(general word processing, complex tables, math, other languages, formatting,
graphics, real-world documents) each with its own subscore, and a **Word** tab
holds the parity target while a **LibreOffice** tab tracks deferred references.
The remaining irreducible floor is pure font-rasterization difference — Chrome
and Word antialias glyphs differently.

```bash
node scripts/parity-parallel.mjs            # full run → parity/out/results.json
DXW_PARITY_FAST=1 node scripts/parity-parallel.mjs   # skip slow appearance passes
node scripts/parity-render-report.mjs       # rebuild parity/out/report.html + report.png
```

See [`docs/EVALS.md`](docs/EVALS.md) for the full methodology.

## How we got here

The certified run measures **1,154 Word-authored pages** across 91 fixtures at a
**mean structural severity of 0.026%**, with exactly two pages above 1% — both a
consequence of a licensed font we can't ship, not a layout bug. Getting there
took a real layout engine (not a docx-to-HTML mapping), a measurement rig that
treats desktop Word's own PDF export as ground truth, and a long campaign of
disproving "it's just a rasterization floor" — complex-script pages that looked
irreducible at 12–17% severity turned out, nearly every time, to be real bugs
(Indic space-font routing 17.26% → 1.30%, Arabic RTL 12.98% → 0.00%, kashida
justification 14.02% → 0.95%).

The full write-up — the architectural bet, the structural-severity metric, the
weird-but-discoverable Word behaviors, and the keystroke-latency work that took
editing from 5 s to ~250 ms — is in
[`docs/PARITY-JOURNEY.md`](docs/PARITY-JOURNEY.md).

## Docs

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — how the parse → layout → render pipeline fits together
- [`docs/DISCOVERIES.md`](docs/DISCOVERIES.md) — ledger of non-obvious Word behaviors we measured (justify pack-vs-break rule, tcW/grid autofit semantics, retina hairline antialiasing, canvas ligature measurement, …) and the probe methodology that established them
- [`docs/EVALS.md`](docs/EVALS.md) — the rendering parity eval: what it measures, the category taxonomy, and how to run it
- [`docs/LIMITATIONS.md`](docs/LIMITATIONS.md) — **Known limitations**: what editing doesn't support yet, shallow toolbar pickers (quick wins), rendering gaps, and untested areas
- [`docs/PARITY-JOURNEY.md`](docs/PARITY-JOURNEY.md) — the long-form story of how we got to 0.026% mean severity: the layout-engine bet, the metric, the war stories, and the editing-perf work

## Roadmap

- [ ] Editing phase 2: caret + text insert/delete, paragraph-level edits (alignment, indents, spacing), then page-layout edits (margins, sections)
- [ ] Floating/anchored DrawingML objects with text wrap (VML lines/textboxes already supported)
- [ ] Footnotes/endnotes, comments, tracked changes
- [ ] Font embedding (`fontTable.xml` + embedded font parts) and font substitution metrics
- [ ] Table autofit (content-based column sizing), vertically merged cell content spanning
- [ ] `keepNext`, right/center/decimal tab alignment with leaders
- [ ] Continuous section breaks sharing a page, vertical page alignment
- [ ] Canvas/SVG renderer, print/PDF export

## License

PolyForm Noncommercial 1.0.0 — free for noncommercial use; commercial use requires permission. See LICENSE.
