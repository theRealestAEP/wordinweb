# DocxInWeb Architecture

A linear walkthrough of the pipeline, with the design decision behind each
stage and the alternative it rejected.

## 0. The framing decision

A `.docx` file contains no positions — Word's layout engine decides where
every line and page break lands. That leaves two architectures:

1. **Translate to flowing HTML/CSS** (docx-preview, mammoth): fast to build,
   but pagination, page numbers, and header geometry are unachievable —
   the browser's layout algorithm is not Word's, and you can't ask it where
   a page ends.
2. **Reimplement the layout engine** and use the browser as a display.

DocxInWeb is architecture 2. Everything follows from one contract: **layout
produces absolutely-positioned primitives; nothing downstream makes a spatial
decision.**

```
bytes → Package → DocxDocument (model + retained XML) → layoutDocument()
      → LaidOutPage[] of {text|rect|edge|image} → renderToDom() → pixels
                                                → DocxEditor (mutates XML) ↩
```

## 1. Container — `src/zip.ts`

OPC (zip) access via `fflate` — the **only runtime dependency** of core
(~8 KB, sync, browser+Node). Synchronous API on purpose: documents are small,
and async would infect every caller. `resolvePartPath()` normalizes OPC's
relative relationship targets so the rest of the code sees canonical part
names.

## 2. XML — `src/xml.ts`

A ~200-line hand-written scanner producing `{name, attrs, children, text}`.

*Why not a library / DOMParser?* OOXML is a narrow dialect: machine-generated,
well-formed, no DTDs, no meaningful mixed content. A general parser pays for
generality that's never used, and DOMParser doesn't exist in Node (tests).

*The namespace simplification:* prefixes are kept verbatim and lookups match
**local names** (`child(el, "p")` matches `w:p`). Formally prefixes are
arbitrary; in practice every producer uses the standard ones. This one choice
deleted an entire namespace-resolution subsystem and is robust to prefix
variation.

`serializeXml`/`cloneXml` support editing write-back: OOXML elements either
have children (containers) or text (`w:t`) — never both — so the writer is
trivial and lossless for our parts.

## 3. Units — `src/units.ts`

OOXML mixes twips (1/20 pt), half-points (font sizes), eighth-points
(border widths), and EMUs (images). **Everything converts to CSS px (96 dpi)
at parse time**; floats are kept unrounded until paint so error never
accumulates.

## 4. Model — `src/model.ts`

Typed IR: `Section → Block (Paragraph|Table) → Run → RunContent`.

Two load-bearing decisions:

- **Sparse property bags.** Every `RunProps`/`ParaProps` field is optional and
  `undefined` means "not set at this level". Word formatting is a cascade
  (docDefaults → style chain → direct); a correct cascade must distinguish
  "explicitly off" from "inherit". Sparse bags + a merge-where-defined-wins
  give the cascade in a few lines.
- **`src` back-references.** Runs/paragraphs/text keep pointers to their
  source `w:r`/`w:p`/`w:t` elements. The model is a *derived view*; the XML
  tree is the document (see §12).

## 5. Properties — `src/parse/properties.ts`

Mechanical `rPr`/`pPr` translation: OOXML on/off semantics (`<w:b/>`=true,
`w:val="0"`=false, absent=undefined), theme colors with tint/shade math,
borders, tabs, and line spacing normalized to one type
(`auto` multiplier | `atLeast` px | `exact` px).

War story: `parseBorder` once ran the color parser against the border
element, whose `val` is the border *style* — producing CSS color `#single`,
which browsers reject silently. Every border in every document was
invisible. Borders read `color`/`themeColor` attributes explicitly now.

## 6. Styles — `src/parse/styles.ts`

`docDefaults` + styles with `basedOn` chains. Resolution walks root-first and
merges; a chain-length guard survives cyclic `basedOn` in corrupt files.
Resolution is **lazy** (functions called at layout time) rather than baked in
at parse time, so edits never invalidate a precomputed cascade.

## 7. Numbering — `src/parse/numbering.ts`

The `numId → abstractNumId → levels` indirection, plus formatters (roman,
letters — `27 → AA`, Word repeats rather than base-26). **Counters do not
live here**: list state depends on document order, so the layout engine owns
counters and this module stays pure. Symbol/Wingdings private-use bullets
(`U+F0B7`) map to Unicode so lists render without legacy fonts.

## 8. Sections & facade — `src/parse/section.ts`, `src/docx.ts`

`sectPr` → page geometry, header/footer refs (default/first/even), `titlePg`,
`pgNumType`, columns. `DocxDocument` assembles the package: main part located
via `_rels/.rels` (never hardcoded), theme → styles → numbering → settings →
body → header/footer parts, each part with its **own** relationship map
(image rels are per-part).

Body parsing honors OOXML's odd section encoding: a `sectPr` inside a
paragraph closes the section *retroactively*; the last `sectPr` sits at body
level. The model normalizes both into clean "sections of blocks".

## 9. Measurement — `src/layout/measure.ts`

`TextMeasurer` is an injected interface:

- `CanvasMeasurer` (browser): `measureText().width` + `fontBoundingBox*`
  metrics — the OS font stack's real numbers, which is what makes line
  heights track Word. Width results memoized by font+string; documents
  repeat words massively.
- `ApproxMeasurer` (Node/tests): a character-width table. Not accurate —
  **deterministic**, so layout tests assert stable numbers on any machine.

Injecting the measurer keeps the engine a pure function of
(document, measurer) — critical for reproducing browser-only bugs in tests.

## 10. Layout — `src/layout/inline.ts`, `src/layout/engine.ts`

**inline.ts** flattens a paragraph into pre-measured *atoms* (word fragments,
spaces, tabs, images, breaks), then breaks lines greedily. Justification
distributes slack across space atoms (never the last line). Right/center tab
stops measure the upcoming text so `text → tab → Page N` right-aligns.
Line height = spacing rule applied over the line's max font metrics, with
text **bottom-anchored** in the line box — a deliberately simple rule that
matches Word closely in all common cases (and exactly for `exact` spacing,
which is why pleading-paper line numbers align with body text). Empty
paragraphs take their height from the paragraph mark's run props.

**engine.ts** walks sections → blocks with a `{page, column, y}` cursor.

The keystone is page creation. Word's rule is
`bodyTop = max(marginTop, headerDistance + headerHeight)` — you must lay out
the header to know where the body starts, but headers contain `PAGE/NUMPAGES`
fields that aren't known yet. Resolution — two passes:

1. At page creation, lay the header/footer out **for measurement only**
   (field text width doesn't change height), set the body box, discard items.
2. After pagination completes (total known), lay headers/footers out again
   per page with exact field values and emit.

That split is the entire mechanism behind a correct "Page 3 of 7".

Also here:
- **Widow/orphan planning** is a simulation pass that computes break indices
  before emission. It must simulate page state — reading live cursor state
  mid-simulation caused an infinite loop once (regression test exists), and a
  progress guard forbids re-adding a break index.
- **Frames**: headers, footers, and table cells all share `layoutFrame()` —
  blocks into a fixed-width, unbounded-height box, offset into place by the
  caller. Tables lay out each cell as a frame; row height per `trHeight`
  rules; `tblHeader` rows repeat after page breaks.
- **Anchored shapes** (VML lines/textboxes — pleading paper): positioned
  against page/margin/text origins without occupying inline space. Inside
  header frames, coordinates are emitted as `pageCoords − frameOrigin` so the
  later frame offset cancels exactly.
- **List counters** live on the engine, incremented in document order with
  deeper levels reset, per-`numId`.

Output: `LaidOutPage[]` of `text` (x, baseline, font, props, src), `rect`,
`edge`, `image`. Renderer-agnostic by design — a canvas/SVG/PDF backend
consumes identical geometry.

## 11. Rendering — `src/render/dom.ts`

Each primitive → one absolutely-positioned element; the browser never
reflows the document. Edges are zero-size divs with a single CSS border
(dashed/dotted/double for free). Zoom is one `transform: scale()` on the page
surface, so item coordinates never change. Pages set
`-webkit-font-smoothing: antialiased` to match Word-on-Mac's grayscale
smoothing (Chrome's subpixel default reads visibly heavier). Every text
element registers an `(element ↔ item)` binding — the bridge editing walks
backwards.

## 12. Editing — `src/edit/`

**The XML tree is the source of truth; the model is never exported.**
Serializing a model back to OOXML silently destroys everything unmodeled
(bookmarks, comments, footnotes, custom XML). Instead, commands mutate the
retained XML tree, `doc.refresh()` re-derives the model, relayout re-renders,
and `doc.save()` re-serializes only `document.xml` + header/footer parts into
the original zip. Tests assert untouched parts round-trip **byte-for-byte**.

Formatting flow: DOM selection → `selectionToSegments` (via render bindings)
→ `(run, w:t, charStart, charEnd)` → `applyRunFormat`: full coverage patches
`rPr` in place; partial coverage splits the `w:r` into before/middle/after
(cloning `rPr` onto each so existing formatting survives), patches the
middle, with `rPr` children inserted in schema order so Word accepts it.

Typing (`DocxEditor`): `caretRangeFromPoint` → `(w:t, offset)`, splice the
string, refresh + full relayout per keystroke — measured in single-digit ms,
so incremental relayout is deferred complexity, not a requirement. The caret
is our own positioned element: contentEditable's editing model cannot be
reconciled with absolutely-positioned spans (the trap that kills most web
docx editors).

## 13. React & packaging

`@docxinweb/react` is a thin wrapper: async source loading, awaiting
`document.fonts.ready` before measuring (measuring before fonts load yields
wrong widths everywhere), lifecycle, and the `editable` flag — default off,
pure viewer; when on, an api (`applyFormat`, `getSelectionFormat`, `save`)
arrives via `onReady`. All logic lives in core so other framework wrappers
stay thin.

npm workspaces monorepo: `packages/core` (zero-config publishable),
`packages/react`, `apps/demo` (Vite; dev-deps and fixture generators never
touch published packages; source aliases give instant HMR).

## The through-line

1. Spatial truth lives in exactly one place — the layout engine.
2. The file's XML is the document; models and pixels are derived views.
3. Write the code the docx dialect needs, not what the general problem
   suggests — hand parser, local-name matching, sync zip, injected measurer.

Speed follows from (3), fidelity from (1), safe editing from (2).
