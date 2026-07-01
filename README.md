# DocxInWeb

Open-source, high-fidelity **Word/.docx viewer (and, soon, editor) for the web** — think Google Docs' rendering quality, embeddable as a single React component.

```tsx
import { DocxView } from "@docxinweb/react";

<DocxView source="/report.docx" zoom={1} />
```

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

## Roadmap

- [ ] Editing: cursor/selection mapping, model mutation + incremental relayout, docx write-back
- [ ] Floating/anchored objects with text wrap
- [ ] Footnotes/endnotes, comments, tracked changes
- [ ] Font embedding (`fontTable.xml` + embedded font parts) and font substitution metrics
- [ ] Table autofit (content-based column sizing), vertically merged cell content spanning
- [ ] `keepNext`, right/center/decimal tab alignment with leaders
- [ ] Continuous section breaks sharing a page, vertical page alignment
- [ ] Canvas/SVG renderer, print/PDF export

## License

MIT
