# Discoveries: how Word actually behaves

A ledger of non-obvious findings made while chasing render parity — each one
cost real investigation time and would be easy to re-litigate later. Every
entry says what the symptom was, what the real cause turned out to be, where
the fix lives, and how the claim was proven (usually a probe document
exported through Word itself; see `scripts/make-*-probe*.py|mjs`).

Rule of thumb from these: **never calibrate against our own measurements or
against pdfminer word extents — build a probe doc, export it through Word,
and read the geometry back out of the PDF.**

---

## Fonts & text measurement

### Office fonts are app-private on macOS
Browsers silently fall back to Helvetica when a file asks for "Calibri"
(~8% wider), drifting line breaks by about a word per line. Fix:
metric-compatible substitutes (Calibri→Carlito, Cambria→Caladea via
@fontsource), a `cssFont` stack in `measure.ts`, and `document.fonts.load`
in `DocxView` (canvas measurement never triggers webfont loads on its own).
Carlito's advance table was verified **byte-identical** to Calibri's
(fontTools over Word's own `DFonts/Calibri.ttf`).

### Word lays out with exact nominal hmtx advances — nothing else
No kerning, no ligatures, no hinting, no quantization. Proven with
`scripts/make-advance-probe.mjs` (20×-repeated-char paragraphs on a huge
landscape page → per-char advances from the PDF match hmtx to PDF write
precision, ±1.6 milli-em). Two consequences:

- **Canvas kerning**: `ctx.fontKerning = "none"` is required (canvas kerns
  by default, ~0.2% narrow).
- **Canvas ligatures**: `fontKerning` does NOT disable ligatures — "ffi" in
  *officia* measured 0.5px narrower than Word while the DOM painted
  ligature-free (`font-variant-ligatures: none`). Fix:
  `ctx.textRendering = "optimizeSpeed"` in `CanvasMeasurer.setFont`.

### Word-mac's PDF export is a lying witness for advances
Justified (and even paragraph-final) lines in Word's PDFs carry per-glyph TJ
adjustments (±5 milli-em) that are writer artifacts, not layout. An entire
earlier justify calibration ("16% accept / 19% reject") was fitted to this
noise. Per-word widths summed from body-text PDFs cannot be trusted below
~0.2pt; isolated left-aligned probe text can.

### Chrome fontBoundingBox integer-rounds
Arial reports 16 where the true hhea value is 16.87 — vertical error
compounds ~1 line per 40. Fix: `WORD_FONT_METRICS` in `measure.ts` = exact
hhea ascender/descender/lineGap read from the font files Word uses.

## Line breaking & justification

### Word's justify packing rule is a pack-vs-break comparison, not a threshold
Word packs one more word onto a justified line **iff the space compression
this needs is at most half the stretch that breaking before the word would
leave behind, capped at 25%** (`JUSTIFY_MAX_COMPRESS` /
`JUSTIFY_STRETCH_FACTOR` in `layout/inline.ts`). A wide word packs at 24%
compression while a narrow one is rejected at 12% — no flat threshold can
model that (the sample lorem paragraph needs accept-at-21.6% AND
reject-at-19.7%). Mapped empirically by `scripts/make-justify-probe*.py`:
sweeps of the needed compression across final words of different widths,
decisions read back from Word's export.

### Words never split at formatting-run boundaries
A word split across `w:r` runs ("(" + "“Cobbery”") is one breaking unit in
Word. The already-placed head backtracks to the next line on break
(`breakParagraph`), instead of leaving a fragment behind.

### Word collapses adjacent paragraph spacing to the larger value
spacing-after + spacing-before of neighboring paragraphs don't add; Word
takes the max (measured from PDF baseline positions). Engine keeps a
`lastParaSpacingAfter` chain, reset at pages/columns/tables.

## Superscript / subscript

### Word's vertAlign geometry (probe-vertalign, measured at 11pt and 22pt)
- scaled size = 65% of base **rounded to half-points** (11pt→7pt,
  22pt→14.5pt — the ratio is not constant!)
- superscript raise = 7/22 of the *unscaled* size (exactly 3.5pt at 11pt)
- subscript drop = 1/11 of the unscaled size

### A baseline shift the renderer never applies is invisible
The engine computed shifted baselines for years while the DOM renderer
bottomed every glyph box on the line box (flex-end + line-height), so
superscript rendered as small text sitting mid-line. Baseline-shifted runs
now carry an explicit glyph box (`glyphTop`/`glyphBoxH` on `TextItem`).

## Tables

### Word ignores authored grids unless cells carry tcW
A `tblGrid` with no `tcW` on any cell is discarded even when it looks
realistic: probe-tablegrid gave an "x" column 5.75pt against its 4680-twip
grid entry. Word writes `tcW` on every cell it lays out, so grid+tcW is
effectively Word's cached layout — trust it; otherwise autofit
(`resolveGridWidths` in `layout/engine.ts`). Corollary: our column drag
must *create* `tcPr/tcW` (as Word does) or the dragged widths get autofit
away on the next layout.

### Hairline borders antialias to gray on retina displays
Word's default 0.5pt table border is 0.67 CSS px; placed at fractional
positions it antialiases across ~1.5 device pixels and reads as light gray,
while Word draws crisp dark hairlines. Fix: snap border width AND position
to whole device pixels in `renderEdge` (`render/dom.ts`).

## Editing UX

### Pleading paper headers have no typeable text
The entire header part is one anchored VML textbox (line numbers + rules);
its host paragraph owns no `w:t`, so no caret target exists and a real page
header was impossible to type. Fix: double-click in the top/bottom margin
*band* enters header/footer mode (Word UX, independent of what text is
near), and a part whose last paragraph has no directly-owned `w:t` gets an
empty run lazily (`hfCaretForBand` in `edit/editor.ts`).

## Tooling gotchas

- **Word-on-mac keeps A4 in PDF export** and docx-lib fixtures default to
  A4 — check `pgSz` before comparing anything against Letter assumptions.
- **Word's sandbox** pops a "Grant File Access" dialog the first time
  AppleScript opens a file at a new path (e.g. a fresh worktree); the
  export then hangs until granted. `scripts/word-parity.sh` callers need to
  expect this once per new directory.
- **pdfminer line bboxes are ink extents**, not advance widths — never
  equate `x1 - x0` of a line with its layout width.
- **Vite dep-cache** can serve stale core after edits (restart with
  `--force`), and worktree-symlinked `node_modules` 403s @fontsource unless
  `server.fs.strict=false` (dev-only, never commit).
- The parity reference PDFs live in gitignored `parity/` — exported ONCE
  per fixture via `scripts/word-parity.sh`, then reused by
  `scripts/compare-linebreaks.mjs` (line breaks) and
  `scripts/parity-compare.mjs` (pixels). Word is never invoked per
  comparison run.
