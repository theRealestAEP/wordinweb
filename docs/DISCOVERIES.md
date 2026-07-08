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

## Math (OMML)

### Word's inline math geometry (measured from parity-math at 11pt)
Cambria Math with letters as Unicode math-italic codepoints; scripts at
**8/11** of the base size; superscript baseline **+4/11 em**; inline
fraction numerator baseline **+6.5/11 em**, denominator **−5.5/11 em**,
with a **0.75/11 em** rule centered **+3.125/11 em** above the baseline
(the math axis); medium spacing around binary operators. Implemented in
`layout/math.ts`; pieces render as baseline-anchored glyph boxes and the
rule as a filled rect. Line-break parity: 0 mismatches vs the reference.

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

### Square wrap clears floats by their wrap distance, nothing more
Text beside a wrapSquare image resumes at exactly image edge + the
anchor's distL/R (parity-wrapmodes: x matches to a hundredth of a point
with dist=0). The engine folds distT/B/L/R into the float record; no
hardcoded padding.

## Editing UX

### Pleading paper headers have no typeable text
The entire header part is one anchored VML textbox (line numbers + rules);
its host paragraph owns no `w:t`, so no caret target exists and a real page
header was impossible to type. Fix: double-click in the top/bottom margin
*band* enters header/footer mode (Word UX, independent of what text is
near), and a part whose last paragraph has no directly-owned `w:t` gets an
empty run lazily (`hfCaretForBand` in `edit/editor.ts`).

## Tooling gotchas

- **Export references with `open`(1) + AppleScript save-as, not
  `open file name`**: `open -g -a "Microsoft Word" file.docx` goes through
  LaunchServices, which blesses the file for Word's sandbox — no
  "Grant File Access" dialog, works even while the screen is locked. Then
  `save as document "<name>.docx" file name <pdf> file format format PDF`.
  Address the document BY NAME: after a force-quit Word restores its old
  session, so `document 1` may be a stale document (we exported seven PDFs
  of the wrong file this way).
- **Never drive Word's dialogs with blind keystrokes**: System Events
  `keystroke` sequences race the dialog; if it closes early the keys type
  into whatever document is frontmost (potentially the user's). If UI
  scripting is unavoidable, verify the sheet exists before every keystroke
  and prefer clicking named buttons.
- **UI scripting requires an unlocked session**: window queries return
  empty and AppleEvents that need a dialog time out (-1712) while the
  screen is locked; plain document AppleEvents keep working. A modal left
  open makes documents refuse `close` with -1708. **LaunchServices `open`
  doesn't deliver documents into a locked session either** — the save-as
  then targets whatever stale document was already open. And sandbox
  grants are inode-scoped: copying a new file over a previously granted
  path does NOT inherit the grant. Net: Word reference exports require an
  unlocked session, full stop — verify the exported PDF's first line
  matches the fixture before trusting it.

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

## Vertical metrics & pagination (2026-07, per-font recalibration)

- **Word quantizes CUMULATIVE baseline positions, not line heights**: the
  n-th baseline of a run of same-height lines sits at round¼(n × rawHeight)
  from the start, so measured gaps alternate around the raw value (Calibri
  11pt: 13.50/13.25pt around raw 13.428; Arial 11pt × 1.15: 14.50/14.75
  around 14.546). A per-line quantized advance drifts a full line over ~50
  rows and flips page breaks. Engine model: the cursor accumulates RAW
  heights; `emitLine` snaps each painted baseline to the quarter-point grid
  (probe-lh2 + sample p2 gap sequence).
- **A trailing page/column break leaves no line behind**: when `w:br
  type="page"` is a paragraph's last content, the pilcrow stays on the SAME
  line as the break on the old page and the new page starts clean at the
  body top (sample/benchmark p2 headings at exactly 72pt). Only content
  after the break moves. We used to emit the paragraph's empty final line
  at the top of the new page — one full line of phantom capacity loss.
- **Widow pull-back cascades into the orphan rule**: 3-line paragraph,
  2 fit — widow control pulls line 2 back, which leaves a lone first line,
  which the orphan rule then pushes too: the whole paragraph moves. An
  if/else widow-orphan implementation misses the cascade (benchmark p2,
  chronology p1).
- **trHeight is a CONTENT-box height** (probe-trheight): Word renders an
  `hRule=atLeast` row at trHeight + top&bottom cell margins + the row's
  border share (~half of each adjoining sz8 border), and an `hRule=exact`
  row at trHeight + top margin only. Google-Docs exports carry fractional
  twip values ("785.92529296875") that must parse as floats.
- **Line-spacing leading may overhang the body bottom**: Word's page-fit
  test only requires the FONT box (baseline + raw descent) to clear the
  body bottom; the extra leading of a spacing multiple (which sits BELOW
  the baseline in the line box) can hang into the margin (msa p2's last
  1.15-spaced line, probe-footerheight). Engine: `LineBox.fitHeight` =
  font-box extent, used by overflow checks; `height` still advances the
  cursor.
- **Header/footer heights include ALL paragraph spacing** — Word does NOT
  suppress the first paragraph's spacing-before in a footer (an empty
  Heading1 lead paragraph costs its full 14pt spacing-before and shrank
  the msa body by exactly one line: 43 vs 44 lines, probe-footerheight).
  Body bottom = min(pageH − marginBottom, pageH − footerDistance −
  footerH).
- **Column balancing before a continuous break** (parity-colbalance): a
  multi-column section followed by a continuous section balances by HEIGHT
  (nine 2-line paragraphs split 5/4 = 10/8 lines, not 9/9): target =
  bandTop + totalStackedHeight/nCols, a line stays in the earlier column
  while its TOP is above the target, the final column uses the true body
  bottom, and the next band resumes below the TALLEST column. No balancing
  happens at document end (parity-columns fills column 1 first).
- **An empty section-break paragraph takes no vertical space**: a
  paragraph whose only role is carrying `pPr/sectPr` renders no mark line
  (columns start exactly one line-advance below the intro). Same family as
  the trailing-page-break rule.
- **The OMML n-ary element is `m:nary`, all lowercase** — `m:nAry` does not
  exist and Word hard-rejects the whole file ("Word experienced an error
  trying to open the file"), unlike most schema slips which it repairs.
  Found by exporting the same formula from LibreOffice (MathML .fodt →
  `soffice --headless --convert-to docx`) and reading LO's OMML — a good
  Word-free way to get canonical OOXML for any construct.
- **A docx without settings.xml opens in Compatibility Mode**, which
  silently FLATTENS math (an nAry degrades to plain runs on save). Probe
  packages must always carry `compatSetting compatibilityMode 15` or their
  verdicts lie.
- **Word-open bisects go through modal error dialogs**: each rejected file
  leaves a dialog that makes every later AppleEvent time out (-1712) and
  queues subsequent opens, which then deliver at random later moments.
  Dismiss via System Events named-button clicks (`click button "No" of
  window 1`) between attempts, and expect force-quit + relaunch to replay
  queued opens.
- **Word math geometry, advanced constructs** (parity-math2 at 11pt): the
  n-ary operator keeps the SURROUNDING font size (math fonts carry a large
  ∑ glyph; ∑ baseline +0.5pt, ∫ -0.5pt); sum-class limits sit at
  +4.25/-2.75pt beside the operator, integral-class at +6.75/-4.5pt with a
  2.2pt slant stagger; the operand follows 2.5pt after the wider limit.
  Matrix rows pitch 12.75pt with their baseline centroid 0.62pt BELOW the
  main baseline; columns gap 12.2pt. Delimiters keep the font size and swap
  in a taller GLYPH VARIANT (the PDF shows "(" at sz 11 spanning a 2x2
  matrix) - we approximate with a paint-time scaleY around the math axis so
  advances stay natural. Binary-operator spacing applies to FULL-SIZE
  content only: scripts, fraction parts and n-ary limits set tight
  ("i=1" under a sum advances glyph-to-glyph). parity-math2's 2 residual
  line diffs are extraction-order artifacts of vertically stacked pieces,
  not geometry (positions verified within ~1pt; the ∑ advance differs
  because STIX Two Math substitutes for Cambria Math).

## Word template rendering (2026-07, header/footer designs + cover letters)

- **Word's built-in h/f templates decode to five constructs**: inline SDTs
  (placeholder prompts), `w:ptab` alignment tabs (Three Columns), pct-width
  tables (Ion Light), and anchored `wps` shapes with theme fills + text
  boxes (Banded/Ion Dark: white title text INVISIBLE without the fill).
  VML fallbacks carry the full geometry as style keys: mso-position-
  horizontal/vertical (alignment), mso-top/left-percent (‰ of page),
  mso-width/height-percent + mso-*-relative (page vs margin), and
  v-text-anchor (text bottom-anchoring inside the box).
- **Template art is a:custGeom freeform paths** (icons, decorative bands) -
  rendered as SVG paths (PathItem). Fills are theme colors with
  lumMod/lumOff/shade/tint transforms (white bg1 x lumMod 85% = the gray
  bands); DrawingML scheme spellings bg1/tx1/bg2/tx2 need aliasing to
  lt1/dk1/lt2/dk2.
- **wp:positionH/V can be wrapped in mc:AlternateContent** (wp14 percent
  offsets with posOffset fallbacks) - child() lookups silently miss them;
  use descendant search.
- **A bare w:trHeight with no w:hRule means AUTO: the value is ignored**
  (cover-letter "Right side layout table": trHeight 10512 but Word
  content-sizes the row; honoring it as atLeast overflows to a phantom
  page 2).
- **contextualSpacing applies inside table cells** (layoutFrame), not just
  body flow - the RECIPIENT/TITLE/ADDRESS block is consecutive Heading2
  paragraphs whose 20pt spacing-after vanishes between same-style
  neighbors.
- **Adjacent same-border paragraphs merge borders** (no rule at the shared
  boundary), which is what makes a run of bordered paragraphs read as one
  box.
- **Office-private theme fonts need substitutes**: Gill Sans MT (cover
  letter/resume templates) falls back to Helvetica silently and every
  measurement is wrong; macOS ships metrically-similar Gill Sans.
- **Office 3D models (am3d) ship their own render**: `am3d:model3d` embeds
  the .glb via r:embed AND an `am3d:raster rName="Office3DRenderer"` poster
  blip - Word's static/PDF output just draws the poster, so painting that
  image at the model extent is exact parity. No 3D engine needed.
- **Word save-as can wedge on documents with 3D content**: an invisible
  modal makes every AppleEvent time out (-1712) and `quit` reports "User
  canceled" - needs a human to dismiss.
- **Word's dashed borders are [3 1] x line width** (dash operator read from
  its PDF export) - CSS `dashed` is much shorter-dashed. renderEdge paints
  dash patterns as repeating gradients to match the rhythm ("For Sale"
  tear-off separators).
- **Playwright element-screenshots of tall absolutely-positioned pages can
  silently omit content below the viewport fold** - the DOM said 14 rail
  items and a title existed while two element screenshots showed blank
  space. Trust DOM measurements or viewport screenshots when they
  disagree.

## Sanitized real-document fixtures (2026-07)

- **scripts/sanitize-docx.py anonymizes a .docx for use as a fixture**: each
  word in w:t/w:delText is replaced by a deterministic same-length,
  same-caps pseudoword (SHA1-seeded), digits remapped, authors → Reviewer,
  core props → Fixture, external hyperlinks → example.com. Structure
  (styles/tables/SDTs/drawings/fields/breaks) is untouched, and the Word
  reference PDF is exported FROM the sanitized file, so line-break parity is
  preserved by construction. Must split on XML entities (&amp; &#8217;)
  before scrambling or it corrupts them.
- **The pleading-paper line-number sidebar breaks the line-break
  comparator**: the VML textbox column of numbers clusters as its own rows,
  so compare-linebreaks.mjs reports hundreds of false mismatches even
  though the render matches Word pixel-for-pixel (7/7 pages verified). Two-
  column/sidebar layouts need the pixel comparator, not the line comparator.
- **Databound content controls are a rendering gap**: the cover-letter
  templates' SDTs (RECIPIENT NAME, [Item], title bands) bind to document
  properties / a custom-XML data store (w:dataBinding w:xpath); Word renders
  the BOUND value while we render the placeholder text in document.xml. Any
  databound-SDT fixture will diverge until we resolve bindings.
- **caretRangeFromPoint jumps to a distant cell on empty-cell clicks**: the
  browser snaps to the nearest TEXT node, and an empty table cell's
  zero-width anchor is not a caret target, so a click resolves into a random
  neighboring cell. Fix: reject caretFromPoint answers whose element rect
  does not vertically contain the click, and score the nearestCaret fallback
  by vertical-dominant 2D distance (dy*3 + dx) so the caret stays in the
  clicked column/cell.
- **A paragraph border reserves vertical space** for its rule + w:space
  above/below the text, pushing the paragraph off its neighbor so the rule
  sits in the gap. Without reserving it, a top border draws INTO the
  previous paragraph (pleading footer: the caption's top border cut through
  the page number). Added to spacingBefore/After in both body and frame
  layout.
- **Exact-height table rows clip overflow** (w:trHeight hRule="exact"):
  Word hides content past the fixed row height rather than spilling it onto
  the page or paginating (the For Sale flyer is one full-page fixed cell -
  typing into it must not push the tear-off tabs off the page). Engine drops
  cell text items whose line starts at/below the row bottom.
