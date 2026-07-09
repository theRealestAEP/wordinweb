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

### A degenerate (ultra-narrow) column gives every inter-word space its own line
When a section's `w:cols/@space` is larger than the text width (e.g.
`w:space="360000"` on a 6-inch column: `sample-files.com-multi-column`), the
computed column width goes negative and Word clamps it to a sliver that holds
one glyph. In that regime Word does NOT let trailing whitespace hang for free —
each inter-word space wraps to **its own line** (measured: a 42-line body
column is ~37 letters + 5 space-lines). We normally trim trailing spaces away
in `flush`, which packs ~one extra letter-line per word onto the page and loses
~12% of the pages (`wild-multicolumn` rendered 41 pages vs Word's 46). Fix
(`layout/inline.ts`): when `availFor(line) < spaceWidth` (column can't hold even
a space) and a space would overflow the current non-empty line, flush the glyph
line and emit the space as its own line (kept, not trimmed). Guarded so normal
columns are untouched (a normal column's `availFor` dwarfs a space width).

### Multi-page column balancing: fresh-page start + measured final band
Word treats a multi-column section that is followed by a continuous break in
two distinct ways, and BOTH matter (`layoutSection` in `layout/engine.ts`):
- **A section that fits one band** balances that band in place — the next
  section resumes on the same page (parity-colbalance, single-band).
- **A section that OVERFLOWS several pages** while sharing a partial page does
  NOT fill the remaining band: Word moves the whole section to a **fresh page**
  (wild-multicolumn's intro page 1 is left empty below the intro; the 2-col
  body starts on page 2 with its Heading2, whose spacing-before is dropped as a
  page-top paragraph). It then flows full columns page by page and balances
  only its FINAL band so the following 1-col section resumes on it.

Both are handled with a real, break-aware **measure pass**: lay the section with
ordinary full-column flow and record where every column of the final page ends
(`balColEnds`); the balance target is `finalBandTop + Σ(colEnds)/nCols` measured
on the REAL final-page content (not a gapless stacked-height estimate, which
fails once keepNext/widow gaps intervene). A final pass restores the pre-section
state (`snapshot`/`restore` of pages, cursor, counters, bookmarks, footnotes,
line-numbering) and re-lays, arming the balance target on the final page only.
The target per column never exceeds a full column, so the final page stays final
and it converges in one balanced pass. This fixed page 2 (63%→24%) and pages
2-12 / 24-29 (≈45%→≈21%, the single-glyph rasterization floor); parity-colbalance
1.24% and parity-columns 2.30% are unaffected.

### Space-before collapses at a page/column top even for SOFT (non-break) flow
This was the "degenerate-column vertical packing" mystery. The sliver column's
per-line advances are NOT wrong — measured against Word's PDF they match to the
quarter-point (13pt Calibri-Bold Heading2 = 18.25pt/line, 11pt body = 15.50pt/
line, 14pt Heading1 = 19.75pt/line). What drifted was the column's *starting*
offset: a `Heading2` (`before=200` twips = 10pt) landing at the TOP of a sliver
column sat its whole 10pt too low, shifting the entire one-glyph column down and
reading as ~65-70% structural on the metric (whose global-offset search caps at
±4px, so a uniform 10pt/13px shift can't register and scores as residual). Word
p23 heading-top = 74.78pt (= the body content top, no space-before); ours was
82.11pt. Rule: **Word suppresses a paragraph's space-before whenever its first
line comes to rest at the top of a page or text column, whether it arrived there
by a hard break OR by ordinary soft flow.** The engine already dropped it on the
explicit-move paths (hard/section breaks via `suppressNextSpaceBefore`, and the
line-0 overflow path which resets `y`), but a heading relocated to a fresh column
top by the **keepLines/keepNext move** (`placeParagraph`, engine.ts) kept its
before — only the keepNext branch had zeroed it. Fix: re-evaluate against the
FINAL cursor just before the before/after collapse — if `y <= bandTop` (empty
column) collapse the before to just the border reserve. Restricted to a GENUINE
top: a later column of the band (`col > 0`) OR a band that itself begins at the
page body top (`bandTop === bodyTop`). A NEW section band that resumes partway
down a page (a 1-col section under the previous section's balanced columns) is
NOT a page top — its leading heading keeps its before (else p30/p31 regress
22%→78%). Net: wild-multicolumn doc mean 15.39% → 2.76%; p23 64.9→0.9, p39
66.1→0.9, and the entire p13-45 ~14-22% band → <1%; parity-colbalance 1.24 and
parity-columns 2.30 unchanged. (Font note: the demo needs Carlito/Caladea — the
metric-compatible Calibri/Cambria substitutes — actually served; a strict Vite
`server.fs` allow-list silently dropped them, substituting a wider fallback that
inflated EVERY page's width-drift and masked this offset as generic noise.)

### A stale page-break space-before drop must NOT leak past a section boundary — the new section's opener uses carry-remainder
The prior diagnosis of the wild-multicolumn p30-32/p46 residuals as a "table
pagination desync" was wrong: the sec2 `LightGrid` table paginates across p30-31
EXACTLY like Word (same 5+2 row split), and every page's CONTENT matches Word.
The residuals are pure vertical drift. The largest, p32 (38%), was a `w:br
type="page"` interaction. wild-multicolumn's sec3 ends with an empty
`<w:p><w:r><w:br type="page"/></w:r></w:p>` (the trailing-break-leaves-no-line
paragraph) and then an empty `sectPr` paragraph that opens sec4 (a 2-col sliver).
The page break armed `suppressNextSpaceBefore`, and `newPage(true)`'s coalesce
started sec4's `Jade 4:` Heading1 on that fresh page WITHOUT clearing the flag —
so its 24pt space-before was fully dropped and the whole one-glyph column sat
~15pt high (measured: our heading ink-top 75.75pt vs Word 90.75pt; pitch matched
to the quarter-point). Word does NOT drop it: a new section's opener follows the
cross-section **carry-remainder** rule (`max(before, carriedAfter) - carriedAfter`
= 24pt − the intro's 10pt after = **14pt**, landing at Word's 90.75pt). The
page-break drop is meant for ordinary post-break flow WITHIN a section, not for a
following section's first paragraph. Fix (`run()` in engine.ts): clear
`suppressNextSpaceBefore` at every section transition (`prevSp !== null`) so the
carry-remainder governs. This is why the old blanket drop looked right at the
sec1→sec2 boundary (p2): there the opener is a Heading2 (before=10pt) after a
10pt-after paragraph, so the remainder is 0 anyway — it only diverged where the
opener's before exceeded the carried after. Net: **p32 38.2% → 1.0%**, doc mean
2.76 → 1.95; no page of p1-29/p33-45 moved (all still <1%), and parity-columns
2.30 / parity-colbalance 1.24 unchanged. Verified by ink-row measurement of the
Word PDF vs our element screenshots at matched 192 dpi.

Still open (p30 22%, p31 12%, p46 19% — the balanced-band RESUME offset, which
lives in the protected two-pass balancing machinery and was NOT safe to touch
here): on the FINAL page of a degenerate 2-col sliver section, the following
1-col section resumes ~4pt too low. The sliver's last line matches Word to the
sub-pixel (sec2's last glyph `e` ink-top 537.6pt in both), but our sec3 `Jade 7:`
Heading1 lands ~3.8pt low (ink-to-ink) and cascades to the whole table; p46's
sec5 body is the same ~4pt low after sec4's balanced band. The inter-section gap
is 24pt for us (sliver after 10 + heading before-remainder 14) vs Word's ~20pt —
and no after/before-collapse configuration nets anything but 24pt, so the excess
is in `balanceMaxY` (the resume = tallest balanced column bottom), i.e. a
balance-TARGET calibration, not a spacing rule. Because the target protects the
24+ clean balanced pages (p2-29/p33-45, all <1%), closing this needs an isolated
2-col-balanced → 1-col-section probe exported through Word to pin the exact
resume height before touching the target — distinct from, and riskier than, the
space-before rule fixed above.

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

### Display equations are `m:oMathPara`, and it changes the LAYOUT, not just alignment (parity2-equations at 11pt)
A bare `m:oMath` is inline; wrapping it in `m:oMathPara` makes it a *display*
equation, and Word renders it very differently (the fixture dropped from
86.8% → 10.4% severity once this was modeled). Measured from Word's export:
- **Centered** on the content column (`m:oMathParaPr/m:jc` default =
  `centerGroup`) regardless of the host paragraph's own alignment. The engine
  centers any line carrying a display-math span in `layout/inline.ts`.
- **Fractions use the FULL base size** for numerator/denominator (not the
  8/11 script scale inline math uses): quadratic numerator baseline
  **+8.25/11 em**, denominator **−7.25/11 em** at 11pt. Display-ness
  propagates through frac num/den, n-ary operands, delimiter/radical content
  and `func`, but NOT into sup/sub scripts (b², xᵏ stay 8/11, text-style).
- **n-ary operators (∑) enlarge and stack their limits above/below** (not
  beside like inline): the grow-variant ∑ glyph is **14.6pt** wide vs the
  ~9.4pt text glyph (≈**1.55×** the font size, since a browser can't reach
  the font's size variants via plain CSS), upper limit **+16.5/11 em**, lower
  **−14/11 em** from the operator baseline, both at 8/11 script size,
  centered on the widest of {operator, upper, lower}. Integral-class
  operators keep limits BESIDE even in display (unchanged).
- **The display-math line height is the math cluster's own ascent+descent —
  NOT that × the paragraph's auto line-spacing multiplier.** The fixture's
  Normal style is `line=259 auto` (1.08×); applying that to the tall math box
  over-inflated every equation line and cascaded the following body text
  several px too low per section. Skipping the multiplier for display-math
  lines (`finishLine`) made the body baselines land within ~2px of Word.
  Rules (fraction bars, radical vinculums) count toward the box's vertical
  extent too — a display radical's vinculum is the topmost ink.
- `m:f/m:fPr/m:type val="noBar"` is a barless stacked fraction (binomial
  coefficient `(n \atop k)` inside big parens) — parse it and skip the rule.

## Tables

### Table-style banding is CONDITIONAL formatting resolved per cell (w:tblStylePr + tblLook)
A styled table (`LightGrid-Accent1` in wild-multicolumn) carries NO direct cell
`shd` — its blue row-banding, bold header/first-column, and thick header
underline all live in the table STYLE's `w:tblStylePr` blocks (`firstRow`,
`band1Horz`/`band2Horz`, `firstCol`, …), gated by the table's `w:tblLook`
(firstRow/firstColumn/noHBand/noVBand flags, either as attributes or the legacy
hex `w:val` bitmask). Parsed into `Style.condFormats` (`parse/styles.ts`),
resolved per cell in `engine.condFor` against tblLook + row/grid-column position
in ECMA-376 precedence (banding < first/last col < first/last row < corners),
and merged UNDER a direct cell shd/border. Banding row index excludes the header
(`rowIdx − (firstRow?1:0)`); even→band1, odd→band2 at `tblStyleRowBandSize`.
Missing this rendered the section-2 table with white rows and no header rule.

### A table row overflowing the body bottom by a bounded amount stays put
Word's page-fit for a table row lets the row's trailing line-leading and its
bottom rule overhang the bottom margin before it moves/splits the row - the same
font-box overhang the body-line fit allows. wild-multicolumn's 5th data row
missed by ~2.4px (an upstream sub-pixel spacing drift left the table ~5px low)
and Word kept it on the page. A small bounded allowance (`ROW_OVERHANG_TOL`,
~3px) in the row-fit check, SUPPRESSED whenever footnotes reserve the bottom
band (the reserve already accounts for it - wild-doerfp), keeps it. Bounded well
under the ~one-line gap that makes Word move a whole row (parity2-nestedtables
moves a 56pt row with 31pt left), so real page breaks are unaffected. NB: the
reference PDFs show no row actually crossing the margin, so this is calibrated
to reclaim our own drift, not a large measured Word tolerance.

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

### wrapText="bothSides" fills BOTH sides of a mid-column float on the same line band
A square/tight float that sits in the MIDDLE of the text column (room on
both sides) makes Word wrap text on the LEFT then continue into the RIGHT
strip at the SAME y, then drop to the next band — a two-segment line, not a
one-sided shrink (parity2-textboxes square box: left col [72,293], right col
[477,540], measured continuous reading order across the two). Modeled by
having the bounds callback (`makeBoundsAt`) subtract each overlapping float
from the band to yield a list of free intervals (`LineBounds.segments`); the
breaker fills interval 0, and on a word that overflows it hops to the next
interval at the same y before breaking the line. Word won't wrap into a strip
narrower than ~40pt (the old one-sided push-below threshold, reused as the
per-segment minimum). Fixed parity2-textboxes p1 44.7%→11.5%.

### A vMerge="restart" cell does NOT inflate its starting row
Its content height spreads over ALL the rows it spans; each row is sized by
its own UNmerged cells, and only if the merged content exceeds the sum of the
spanned rows is the deficit added to the LAST spanned row (parity2-
nestedtables: "vMerge start (tall)" is 2 lines but rows A and B stay one line
each — 27pt, not 41+27). `computeRowHeights` runs a pre-pass over all
laid-out rows; the restart cell then PAINTS (shading/borders/vertical-align)
at its full spanned height via `spanHeight`, while continue cells only carry
the vertical rules through.

### The mandatory empty paragraph after a table collapses to zero height
OOXML requires a `w:p` after every table (and before a cell/frame end); when
that trailing paragraph is empty Word gives it NO line — a nested table sits
flush against the cell bottom, not a blank line below it (parity2-
nestedtables: the `<w:p/>` after both the L3 and L2 tables cost a phantom
~22.5pt line each — 8pt after + 14.5pt line — and cascaded the whole outer
row ~45pt too tall). `layoutFrame` skips an empty paragraph whose previous
block is a table (anchor-carrying paragraphs are NOT empty — collapsing them
would drop the float). This single rule took nestedtables p1 34.3%→5.3% and
p2/p3 from 10.7/16.0 to 2.9/3.0.

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
- **…and its spacing-AFTER stays on the old page too**: the pilcrow-on-the-
  old-page rule means a break-terminated paragraph's `w:after` must NOT push
  the fresh page's first content down (`placeParagraph`: skip `y += spacingAfter`
  and zero `lastParaSpacingAfter` when the last line has a `forcedBreakAfter`).
  wild-multicolumn's section 2 ends with an empty `<w:p><w:br type="page"/></w:p>`;
  its ~13px spacing-after was landing at the TOP of the next page, which the
  following continuous multi-column section then misread as a shared partial
  page (`this.y > bodyTop`) and skipped to a blank next page — one spurious page
  (47 vs Word's 46). Two guards make this robust: `layoutSection`'s
  `sharedPartialPage` now also requires the page to actually hold content
  (`this.cur.items.length > 0`), and `newPage`'s empty-page COALESCE (pop the
  blank fresh page and start the section on it) now covers `continuous`
  sections, not just nextPage — a continuous section reached over a hard page
  break starts ON that fresh page.
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
- **Header/footer references inherit per-type from the previous section**
  (ECMA-376 §17.10.1): a section that omits a `headerReference`/
  `footerReference` of a given type (default/first/even) uses the previous
  section's reference of that type — Word only blanks a header by pointing
  at an EMPTY header part, never by omitting the reference. wild-athabasca's
  body section carries only a footer, yet Word keeps the previous section's
  running header on all 19 pages (and the earlier TOC section keeps the
  roman-numeral footer it likewise inherits). Missing this dropped the
  header on p13–31, freeing ~1 line/page and losing the 31st page. Fix:
  `inheritHeaderFooterRefs` pass in `parse/document.ts` after section
  assembly (forward, so inheritance chains section→section).
- **keepNext binds transitively — the whole CHAIN moves as a unit**: a run
  of consecutive keepNext paragraphs (heading + sub-headings, or a document
  that styles body paragraphs as headings — heading styles carry keepNext/
  keepLines in styles.xml, not inline) must all land on one page together
  with the first line(s) of the terminating non-keepNext block. Each hop can
  individually fit while the accumulated chain does not, so a single-hop
  check under-breaks (wild-athabasca: a 7-paragraph Heading2/3 chain leaves
  ~12 blank lines at a page bottom in Word; packing it lost a page). The
  engine walks the whole forward keepNext run in `placeParagraph`, sums the
  full heights of the keepNext members plus the terminator's first (+orphan)
  line, and moves the lot when it won't fit but fits on a fresh page.
- **w:beforeAutospacing / afterAutospacing insert one blank line, ignoring
  the literal before/after** (HTML/web-pasted content, `NormalWeb`): Word
  discards the 5pt `w:before`/`w:after` and uses ~one SINGLE line height of
  space (wild-athabasca title page: NormalWeb blocks sit a full line apart,
  27.8pt gaps = single line 13.8 + ~14pt auto). Use the line's
  `naturalHeight` (the line-spacing multiple, e.g. line=480 double, must NOT
  inflate the auto gap). **A trailing EMPTY autospacing paragraph's
  after does not carry across a section break** into the next section's first
  paragraph — else it eats the following Heading1's spacing-before (athabasca
  p6: section-1 Tadulobo needs its full before=24pt, not before−prevAuto).
  The cross-section spacing carry (`lastParaSpacingAfter` across the break)
  zeroes when the previous section's last laid-out paragraph was empty
  (`lastParaWasEmpty`); non-empty content still carries (parity2-sections).
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
- **Vector drawing groups (icons/logos) need a hit overlay to be
  draggable**: their pieces render as separate SVG paths/images with no
  single target. The engine emits a transparent DrawingHitItem over each
  inline group (carrying the source w:drawing); the renderer materializes it
  only in interactive mode; dragging it re-anchors the drawing to the drop
  position (moveDrawingTo), same as an inline image.
- **Header/footer single-click gating must distinguish on-glyph from
  whitespace clicks**: clicking squarely on header text stays gated (inert
  until double-click), but a whitespace click near a page edge whose nearest
  text is an hf span retries the nearest BODY text. Keyed on whether
  caretFromPoint (a real glyph hit) returned a result.
- **Databound content controls render from their CACHED text**, which Word
  keeps in sync with the binding — so real documents render RECIPIENT NAME /
  the author name correctly with no special handling. The cover-letter
  divergence was only in the SANITIZED fixture: scrambling the cached SDT
  text while Word re-resolves the w:dataBinding xpath against the scrubbed
  core.xml. Live xpath resolution (core.xml / customXml) is a future
  robustness nicety for stale-cache documents, not a real-doc bug.
- **w:position (raised/lowered text) grows the line box by the FULL shift,
  additively after the line-spacing multiplier**: a +6pt raise on a Multiple
  1.08 / after-8pt line advances the pitch by exactly 6.00pt (charstyles
  probe, pdfminer baselines: 22.5pt normal pitch → 28.5pt), not 6×1.08. The
  raised run's extension goes above the baseline (baseline moves down within
  the line); a lowered run's goes below. Line-rule "exact" suppresses it.
- **Word's PDF export triple-draws emboss/imprint runs** (three offset
  copies per glyph — pdfminer shows "TTThhheee"); the visible ghost is a
  gray copy offset down-right (emboss) or up-left (imprint). Approximated
  with a 1px gray text-shadow. w:outline paints hairline-stroked hollow
  glyphs (transparent fill, ~0.75pt stroke).
- **VML textpath (WordArt watermark) fitshape geometry**: Word fills the
  shape box HEIGHT with the em (~0.86× box height; the glyph band sits
  slightly above the box's vertical center) and squashes/stretches the
  glyphs horizontally to the box width — not width-fit-at-natural-aspect.
  Also: every PageItem kind must be handled in offsetItem — the wordart
  case was missing, so header watermarks lost the frame's page offset and
  painted exactly marginLeft/headerDistance up-left of Word on every page.
- **Literal TAB characters inside w:t render as real tab stops** (generator
  files; Word normalizes them to w:tab on save). Split into tab atoms at
  atom-build time, keeping the model text (and editing offsets) intact.
- **Word pushes a paragraph below a following paragraph's topAndBottom
  float** (parity2-textboxes): the box anchored at para N's top (posOffset
  0) is positioned from para N's UNDISPLACED position, then earlier lines
  that overlap the band (the section heading, para N−1) reflow below the
  box while the box keeps its first-pass position. Needs iterative/lookahead
  anchor placement — open problem, parked (~8pp on one fixture page).
- **Word's autofit column widths don't fit content+constant-margin**
  (parity-tables): measured slack over PDF glyph advances varies per column
  (10.4px vs 15.9px on sibling columns, no tblCellMar). tblW w:w="100%"
  (invalid per ST_TblWidth) is ignored by Word → table autofits to ~601px,
  not the 624px content width. Open calibration problem, parked.
