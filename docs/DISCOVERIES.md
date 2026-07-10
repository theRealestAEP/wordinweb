# Discoveries: how Word actually behaves

A ledger of non-obvious findings made while chasing render parity — each one
cost real investigation time and would be easy to re-litigate later. Every
entry says what the symptom was, what the real cause turned out to be, where
the fix lives, and how the claim was proven (usually a probe document
exported through Word itself; see `scripts/make-*-probe*.py|mjs`).

Rule of thumb from these: **never calibrate against our own measurements or
against pdfminer word extents — build a probe doc, export it through Word,
and read the geometry back out of the PDF.**

### The Times probe VALIDATES `WORD_FONT_METRICS['times new roman']` — the wild2 drifts are NOT TNR pitch (2026-07)
`scripts/make-times-probe.py` (TNR + bare "Times" at 10/10.5/11/12pt ×
single/double/1.08×/atLeast, each block 50 forced-break lines in one paragraph)
was finally exported through an unlocked Word (`parity/probe-times-word.pdf`)
and regressed (`scripts/read-times-probe.py`, now recursing into pdfminer's
nested `LTTextBox` containers — it previously only saw top-level lines and
found none). A linear fit of baseline-y vs line-index nails the raw per-line
advance immune to Word's quarter-point cumulative-baseline quantization:

| block | config | advance | per-em |
|---|---|---|---|
| A | TNR 12pt single | 13.8029pt | 1.150240 |
| B | TNR 11pt single | 12.6523pt | 1.150211 |
| C | TNR 10pt single | 11.5035pt | 1.150346 |
| D | TNR 10.5pt single | 12.0774pt | 1.150229 |
| E | **Times** 12pt single | 13.8029pt | 1.150240 |
| F | **Times** 11pt single | 12.6523pt | 1.150211 |
| G/H | TNR 11/12pt double(480) | — | 2.30048 (= 2×single, exact) |
| I | TNR 12pt ×1.079(259) | 14.8956pt | 1.241300 (pred 1.15025×1.079167 = 1.241318 ✓) |
| J | TNR 10.5pt atLeast348 | 17.405pt | pins at 17.40pt (= 348tw) ✓ |

- **Word's true single-spaced TNR/Times per-em is ~1.15025** (mean of the six
  50-line fits 1.150246, spread ±0.00007). Our baked total is
  `0.891113+0.216309+0.04248 = 1.149902`. **Delta = +0.000344 em = 0.0041pt/line
  at 12pt.** That is BELOW Word's own quarter-point quantization floor (~60
  lines of TNR-12 to accumulate a single 0.25pt tick), so it cannot flip
  pagination on any TNR-body doc; the earlier "body pitch is correct, do not
  tweak the hhea value" note (below) stands — my 13.803pt merely lands at the
  top of that note's measured 13.78–13.80 band. **No source change made:** a
  global-metric churn for a sub-quantization 0.03 % delta with zero demonstrable
  fixture benefit is exactly the regression risk to avoid. Recorded here so a
  future session with a *demonstrated* TNR-pagination-drift fixture can apply
  gap `0.04248 → 0.042828` (total 1.15025) with confidence.
- **Bare "Times" === Times New Roman EXACTLY** (E/F advances identical to A/B to
  the milli-point). Confirms the `times → Times New Roman` entry now in both
  `METRIC_SUBSTITUTES` and `WORD_FONT_METRICS` (measure.ts) is the correct
  resolution, not merely a good-enough stand-in.
- **The wild2 "Times-family drift" fixtures do not actually hinge on TNR pitch.**
  `wild2-legal-nih-contract` is **Calibri-only** (0 "Times" occurrences in
  document.xml; its −2p drift is the schedule-table row-height deficit, a
  separate probe). `wild2-lit-yiddish-rtl` is **Tahoma-cs-dominant** (Hebrew
  runs inherit docDefault `w:cs="Tahoma"`; a single style declares TNR) — its
  73.9 severity is RTL/Tahoma-line-box driven, not TNR advance. So item 1 could
  never have moved either by touching the Times metric.
- **Reader fix kept:** `read-times-probe.py` now recurses containers and skips
  pages with <3 lines (double/1.08×/atLeast blocks legitimately spill a sparse
  tail line onto a 2nd page).

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

### A balanced band resumes at the balance TARGET, not the final column's raw cursor (the trailing spacing-after must not double-count)
The balanced-band RESUME offset (wild-multicolumn p30/p31/p46, ~5pt too low on
the FINAL page of a degenerate 2-col sliver section). The sliver's last line
matches Word to the sub-pixel, but the following 1-col section resumed ~5pt low
and cascaded into whatever followed (p30/p31 a LightGrid table; p46 an
empty-para + body). The fix lives in the resume line of `run()` (engine.ts),
which was `this.y = Math.max(this.y, this.balanceMaxY)`.

Instrumented, the three cases (px; ÷1.3333 = pt) were:

| case | this.y (final col, incl. trailing after) | balanceMaxY (tallest NON-final col) | target = bandTop + Σ(colEnds−bandTop)/nCols | Word resumes at |
| --- | --- | --- | --- | --- |
| wild p30 | 747.60 | 734.26 | **741.00** | 741 (target) |
| wild p46 | 438.76 | 425.43 | **432.00** | 432 (target) |
| parity-colbalance | 310.47 | **356.94** | 333.57 | 356.94 (balanceMaxY) |

`max(this.y, balanceMaxY)` gives 747.60/438.76 for wild (≈5pt low) yet the
correct 356.94 for colbalance. The final column's raw cursor `this.y` carries
the SECTION's trailing paragraph spacing-after (10pt here), which Word does NOT
bake into the band height — it applies that after via the section-boundary
before/after collapse against the next paragraph's before. Baked into `this.y`
AND distributed once into the target, the after is counted ~1.5×; the visible
excess is `after − after/nCols` = 5pt (10 − 10/2). Word instead resumes at the
balance TARGET (the even column height), which spreads the trailing after by
1/nCols. Rule (`run()`):

    this.y = Math.max(this.balanceMaxY, this.balanceBottom /*=target*/,
                      this.y - this.lastParaSpacingAfter);

- **target** wins for wild (even columns: the balanced columns fell one glyph
  short of the target, the raw cursor overshot by the after);
- **balanceMaxY** wins for parity-colbalance (uneven 5/4 split: the taller
  column's internal after is genuine column height because content follows it in
  the next column — that after Word DOES keep);
- **this.y − trailingAfter** guards the case where the final column overran the
  target with real content (its content bottom then governs).

Net: p46 19.45%→1.53% (exact — the resume Body ink went 356.80pt→351.9pt vs
Word 351.87); p30 22.36%→19.17%, p31 12.08% (the resume/heading is now correct,
p30 align 4px→1px, our `Jade 7:` Heading1 ink 572.7pt vs Word 572.92 — see next
note for their remaining table residual). parity-colbalance 1.24% and
parity-columns 2.30% unchanged; every wild p1-29/p32-45 unmoved (max 1.84%).
Derived with `scripts/make-colresume-probe.py` (2-col-balanced → 1-col-section,
uniform 11pt so line-pitch cancels the leading offset; sliver + normal-column +
after/before-sweep variants) plus the pre-existing `parity-colbalance` Word PDF
as the uneven-column ground truth (local Word could not export a fresh
reference this session, so the wild-multicolumn and parity-colbalance Word PDFs
were the arbiters, cross-checked against the engine's own instrumented values).

### wild-multicolumn p30/p31 have a SECOND residual — the LightGrid-Accent1 table, not the resume
With the resume corrected, p30's 19% / p31's 12% (raw only 3.5% / 1.2%) are the
one LightGrid-Accent1 table (spanning p30→p31), NOT vertical drift of the
section. Two independent gaps, both distinct from the balance resume:
1. **Conditional bold is resolved but never painted.** `condFor` correctly
   layers the `firstRow`/`firstCol` bands' `<w:b/>` into `TableCondFormat.bold`,
   but `paintRow` consumes only `cond.shd`/`cond.borders` — the bold never
   reaches the cell text (which is laid out in `layoutRow`→`layoutFrame` with the
   base run props, before `condFor` runs). Post-painting the weight onto the laid
   glyphs (flip `it.font.bold`) actually REGRESSED the metric (p30 19.17→19.74),
   so Word's rendered weight / which bands bold differs from the naive
   firstRow+firstCol read — needs a fresh Word reference to pin before wiring the
   conditional rPr into cell layout (bold changes glyph advance, so it must feed
   line-breaking, not just paint).
2. **Header row height.** Word's header row pitch is 15.50pt vs our uniform
   14.43pt (~1pt short); the body-row pitches match (~14.5). That ~1pt plus a
   ~1pt table-start gap cascades every data row ~2pt high — a systematic
   misalignment the NCC structural metric amplifies. A LightGrid header
   row-height detail, unvalidatable without Word. Both are table-styling issues
   confined to the single wild-multicolumn table (benchmark/sample/
   parity-columns/parity-colbalance have no tables), so out of the resume scope.

**PROBED 2026-07 (`scripts/make-lightgrid-probe.py`, LightGrid-Accent1, tblLook
04A0, 8×4, faithful reuse of the fixture's styles.xml + theme1.xml):**
- **Bold bands = firstRow ∪ firstCol, EXACTLY — the naive read was right.** In
  `parity/probe-lightgrid-word.pdf` every row-0 cell AND every col-0 cell renders
  `Calibri-Bold` (glyph width 49.51 vs 48.76 for regular); all interior cells are
  regular. (This theme is inverted — major latin = Calibri, minor = Cambria — and
  Normal pins Calibri, so header = Calibri-Bold not a serif.) Our engine paints
  **nothing** bold (DOM fontWeight 400 everywhere). So the p30 regression from
  post-painting was NOT a wrong band set; it was that bold widens the glyph
  advance and must feed **line-breaking**. Fix: thread the resolved conditional
  `bold` (and its rPr) into `layoutRow`→`layoutFrame` so cell text is measured
  bold, then painted bold — not flipped onto already-laid glyphs.
- **Header pitch deficit = a missing CONDITIONAL border in `rowBorderShare`.**
  Word row heights (border-to-border, from the PDF rules): header row = **15.50pt**
  (top rule 1.0pt sz8 at y693.47 → header/row1 rule **2.25pt sz18** at y677.97),
  body rows = 14.43–14.5pt (sz8 1.0pt boundaries). Ours is a UNIFORM 14.43pt
  (= single 11pt Calibri 13.43 + sz8 share 1.0) — body matches, header is
  ~1.07pt short. The firstRow tblStylePr gives the header a **sz18 (2.25pt)
  bottom border**; `rowBorderShare` only reads `tbl.props.borders` /
  `cell.props.borders`, never the tblStylePr conditional borders, so the header
  boundary is counted as sz8 not sz18. Counting it adds (2.25−1.0)/2 = 0.625pt
  (14.43→15.06); the residual ~0.44pt is secondary (thick-border row rounding /
  bold line box). Fix: `rowBorderShare` (and the boundary() helper) must fold in
  `condFor`'s conditional borders per boundary.
- **Neither fix shipped** — both thread conditional formatting into the layout /
  row-height path and touch every style-conditional table, so they need the table
  fixture gate (staging-*/parity-tables/parity2-nestedtables) + line-break suite
  before landing. The probe pins the targets precisely (bold set confirmed; header
  = +0.625pt conditional-border share) so the wiring can be done and validated.

### firstLine and hanging share ONE mutually-exclusive slot across the style cascade
`w:ind/@firstLine` and `@hanging` are the same first-line-indent property (they
can't both apply). When a derived style's `w:ind` specifies one, it must CLEAR an
inherited value of the other — otherwise a parent style's `hanging` leaks past a
child's `firstLine="0"`. On a CENTERED paragraph the leaked hanging shifts the
line left by HALF the hanging (a `hanging=360`tw=18pt centered heading landed 9pt
left of its own body lines: gatech `MainBodyHeadings` firstLine=0 under
`CoverPageSingleSpace` hanging=360). Fix (`parse/properties.ts`, in
`parseParaProps`): when a `w:ind` sets `hanging`, also emit `indentFirstLine=0`,
and vice-versa, so the shallow style merge overrides the sibling attribute.
`hanging` wins when both are (invalidly) present. Left/right indents still merge
attribute-by-attribute (unchanged). Fixed gatech p2 14.5→2.7, p5 5.5→1.9.

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

### A trusted fixed-unit grid that overflows the column is NOT scaled down
Word honors a Word-authored grid (every cell carries `tcW`) even when the table's
authored width exceeds the content column — it keeps the column widths and lets
the table HANG into the right margin, rather than shrinking to fit. Our
`resolveGrid` used to scale any `total > available` grid down to the column, which
shifted every row left by a fraction of the overflow (gatech TOC 2-col table,
grid 9129tw in an 8640tw column, shifted ~4.6pt/92tw left; the right-aligned
marker cell AND the text cell both slid). Fix (`resolveGridWidths`): in the
trusted-grid branch, a non-percentage (`dxa`/absent) grid whose `gridTotal >
available` returns the raw `tbl.grid` unscaled. Percentage widths are relative to
the column, so they still fit. Fixed gatech p11 14.1→3.2, p10 7.3→5.4.

### A nested table that overruns its host CELL is autofit (confined), not hung
The "trusted overflow hangs" rule above is a BODY-level rule. A nested table
(one laid inside a table cell) that overruns its host cell is instead confined
by autofit — Word never lets a nested table hang past its parent cell border
(staging-grid4: L1>L2>L3>L4>L5, each level's authored grid overruns its host
cell yet Word keeps every level inside the 200pt middle column). Two bugs
compounded here: (1) nested tables went through `resolveGrid` (uniform down-
scale) not `resolveGridWidths`, and each level re-scaled its ALREADY-scaled
parent cell, so L5 collapsed to a ~6pt sliver stacking one glyph per line;
(2) autofit's per-column min/pref width ignored nested-table cells (only
paragraphs were measured), so a cell hosting a table contributed no width demand.
Fix (`layout/engine.ts`): `layoutTableInFrame` calls `resolveGridWidths(tbl,
width, nested=true)`; the trusted-overflow branch, when `nested`, falls through
to autofit instead of returning the raw grid; and a new `columnMinPref` /
`measureTableWidths` pair recurses so a cell's nested table contributes its own
min/pref total (the innermost L5's min-width bubbles up and widens every
enclosing "holds L…" column). staging-grid4 p1 43.5→27, tblextreme p1 improved;
parity2-nestedtables unchanged (3.1). Residual is Word's proprietary autofit
DISTRIBUTION between columns (same uncracked calibration as parity-tables) — our
share over-weights the deep-nest column vs the "side A/B" column.

### A trailing COLUMN break leaves the pilcrow on the NEW column (page breaks don't)
The "trailing page/column break leaves no line" rule is TRUE for page breaks but
NOT for column breaks. A paragraph ending in `w:br type="column"` puts its
paragraph mark + spacing-after at the TOP of the next column/page, so the
FOLLOWING paragraph starts one empty line + after lower — whereas a trailing
page break keeps the pilcrow on the old page and the next page starts clean at
the body top (staging-breaks: "Forced into column two" lands at ink-top 96.78pt
= body-top 74.28pt + one empty Normal line 14.5pt + after 8pt = 22.5pt below a
page-break paragraph's clean top). Fix (`layout/inline.ts`): a trailing column
break flushes the text line with `forcedBreakAfter="column"` then an empty
pilcrow line on the new column; a page break keeps the old single-line flush.
Only staging-breaks uses explicit `w:br type="column"` (column BALANCING is a
different path), so parity-columns/colbalance/wild-multicolumn are untouched.
Fixed staging-breaks p5 98.5→0.7, p6 100→0.6.

### Consecutive page/column breaks in one paragraph make a BLANK page between
`<w:p>{break}{break}text</w:p>` = the empty region between the two breaks is its
own blank page (staging-breaks: two `w:br type="page"` produce a blank page 2
between "Before the breaks." and "After two…", matching Word's 7-page count vs
our old 6). The FIRST leading break is a break-BEFORE consumed by `placeParagraph`
at the block level (drops spacing-before, emits no mark line); a SECOND/later
consecutive leading break must emit an empty break line to advance past the blank
page. Fix (`layout/inline.ts`): the leading-break skip only fires once
(`consumedLeadingBreak`); subsequent leading breaks flush an empty break line.
This single page-count fix took staging-breaks mean 71.8→~6.

### w:framePr positioned text frames are absolute-positioned floats
A `w:framePr` with a width (not a dropCap) lifts the paragraph out of normal flow
and paints it at an absolute anchor position while body text wraps around it
(staging-frames). Parsed into `ParagraphProps.frame` (`parse/properties.ts`);
`placeFrameParagraph` (`layout/engine.ts`) resolves the origin from hAnchor/
vAnchor + x/y (page→x, margin→marginLeft/Top+x, text/column→colX/cursor+x;
measured exact against Word), lays the paragraph at the frame width, and registers
a wrap float (wrap=around→square bothSides). It does NOT advance the cursor or the
spacing chain. Numeric x wins over a named xAlign. Fixed staging-frames p2
7.2→2.1. OPEN: a frame anchored BELOW earlier body content reflows that PRECEDING
content around itself (staging-frames' page-anchored top-right box overlaps the
document's opening Heading1, forcing it to wrap to 2 lines and pushing the whole
body down ~23pt) — this is the same parked "lookahead anchor placement" problem
as parity2-textboxes; without it staging-frames p1 stays ~87% (a pure vertical
cascade from the 1-line-vs-2-line heading). NB: "Calibri Light" (Heading fonts)
has no metric-compatible substitute — it maps to Carlito (= Calibri REGULAR), so
long headings that Word wraps via the wider Calibri-Light advances may fit one
line for us even absent the anchor issue.

### A full-width `wrap="notBeside"` frame in a multi-column section is a banner that spans ALL columns
IEEE two-column papers put the title/authors in `w:framePr wrap="notBeside"`
paragraphs (styled `Title`/`Authors`) whose width (from the STYLE: `w:w=9360`,
`hAnchor=page`, `xAlign=center`) is the full page text width, wider than one
column. Word lays these ACROSS both columns at the top of the section and starts
the two-column body BELOW them; the engine was flowing them as ordinary column-0
paragraphs, so the title overlapped the body. Fix: `placeBannerFrame`
(`layout/engine.ts`) — a `notBeside` frame whose `w > colWidth` in a >1-column
section lays full width, stacks with adjacent banner frames (consecutive
same-signature frames = one logical Word frame, no vSpace between their lines),
and pushes the column band (`bandTop`) below itself via `flushBannerBand`.
- **framePr merges attribute-by-attribute across the style cascade.** The IEEE
  Authors paragraph's DIRECT framePr is `<w:framePr w:h="781" w:hRule="exact"
  w:x="1569" w:y="-213" w:wrap="notBeside"/>` — it has NO width, so it must
  inherit the Authors STYLE's width/anchor/xAlign while overriding h/hRule/x/y.
  The parser now emits only the attributes actually present (all frame fields
  optional in the model) and `mergeParaProps` deep-merges `frame`
  (`{...base.frame, ...over.frame}`); `Engine.resolveFrame` fills defaults at
  layout time (a widthless notBeside frame defaults to the full text width). The
  exact `w:h=781` (~52px) reserve is what lands the column band at Word's height.
- OPEN (IEEE p2-4, still ~72%): the body columns' line pitch matches Word to
  ~0.02pt and the banner now spans correctly, but the 2-col body still drifts —
  the figures (e.g. a magnetization chart anchored in the right column) are not
  positioned like Word, and the authors frame's `w:y=-213` raise vs Word's cached
  position needs a Word-export probe. Banner fix alone leaves the doc mean ~flat
  (the 4 pages are dominated by figure/column drift, not the title band).

### Space-before at a page top after a page break is a compatibilityMode-15 behavior
Word 2013 (`w:compat` `compatibilityMode="15"`) suppresses a paragraph's
space-before when it lands at the top of a page after a page break; Word 2010 and
earlier (mode <= 14) KEEP it. `wild2-med-nccih-protocol` is mode 14: a `Heading1`
opening with a leading `w:br type="page"` (before=18pt) sits 18pt below the top
margin, and a `Heading2` reached by a trailing page break sits at margin+before —
NOT flush at the margin. The engine hard-dropped both. Fix: parse
`compatibilityMode` (`docx.ts`); gate the leading-break drop and the
`suppressNextSpaceBefore` (trailing-break) drop on `compatibilityMode >= 15`
(`layout/engine.ts`, `placeParagraph`). Two subtleties held it together:
- **The mode-14 "keep" applies ONLY to explicit page breaks, not pure soft
  overflow.** A heading that soft-overflows to a page top still collapses its
  before in all modes (`atPageOrColumnTop` unchanged for the soft path). And a
  COLUMN top (col > 0) still collapses in all modes — `wild-multicolumn` is ALSO
  mode 14 yet its sliver-column Heading2 drops its before; that path is untouched.
- Fixed nccih p3 81%→1.5% and lowered p2/p5/p8; doc mean 24.3→20.8. Residual:
  the title page (p1) and a section-boundary case where Word adds ~21pt above a
  continuous-section Heading2 that we don't reproduce (p4/p9 — the empty
  page-break section-break paragraph's contribution; needs a Word probe). A
  side effect is a pagination micro-shift on p9 (3%→67%) that resynchronises by
  p10 — the doc's vertical model is now correct at p3 but not yet at that section
  boundary, so the accumulated content lands one band off for one page.

### A continuous section that RESTARTS the page count (same format) still shares the page — only a FORMAT change forces a new page
`wild2-legal-ca-agreement`'s schedule sections are `type="continuous"` with
`<w:pgNumType w:start="1"/>` (restart the decimal count) and a different footer.
Word flows them onto the SAME page as the preceding section's content (the shared
page keeps its own number; the restart takes effect on the section's next full
page). The engine promoted ANY continuous section that set `pageNumberStart` to a
page break (`layout/engine.ts`, `canContinue`), inserting a spurious extra page
after section 3's `5.2` and pushing every downstream page off by one (24 pages vs
Word's 23). The rule that motivated the promotion (wild-gatech) is really a
FORMAT change: gatech's front-matter section is `pgNumType fmt="lowerRoman"
start="4"` — decimal→roman can't coexist on one sheet. Fix: promote only when the
page-number FORMAT differs, not when the count alone restarts. ca-agreement 24→23
pages, mean 55.3→33.5; gatech unchanged (1.99). Residual p11-20 drift is a
separate schedule-content density issue.

### PARKED without Word exports (2026-07): wild2-med-phase23-protocol TOC line count
The 70-page phase23 protocol renders 72 pages because its `TOC` field content
occupies ~2 more pages than Word's (our TOC spans p7-9; Word's is p7 only, p8 is
already body "9.6.4"). The divergence is in how many TOC entry lines / page-number
column widths we emit vs Word's cached TOC — every body page is then offset by the
front-matter delta. Cracking it needs a Word-export probe of the TOC field
geometry (entry wrapping, PAGEREF page-number widths, tab-leader fill), which the
screen-locked session forbids. Left untouched.

**PROBED 2026-07 (`scripts/make-toc-probe.py`): the TOC layout PRIMITIVES are
correct — the phase23 residual is fixture-specific, not a TOC-rendering bug.**
A 40-entry probe reusing phase23's real styles.xml + theme1.xml (TOC1/2/3 with
their `right dot-leader @9350tw` tab, Hyperlink, docDefaults, docGrid
linePitch=360) exported to `parity/probe-toc-word.pdf`. Word laid it out in
**43 visual lines on 1 page** with exactly two entries wrapping (the two >70-char
titles); pitch alternates ~13.0pt (TOC1, `line=240` single) / ~15.5pt (TOC2/3,
inherited `line=276` ×1.15). Our engine on the same probe (DOM, :5317) produces
**the identical 43 lines, the SAME two wrap points, and the same pitch
alternation** — the dot-leader fills one span to the page number (NOT wrapping
to extra dot rows, an earlier grouping-artifact worry), TOC1's single-spacing
override is honored, and wrap columns match. So dot leaders / entry wrapping /
per-style pitch are all correct.
- The real fixture (90 TOC entries: 12 TOC1 / 35 TOC2 / 43 TOC3) still renders
  its TOC over OUR pages 7-9 (42+39+11 lines, 72 total pages) as before. Because
  the faithful probe matches Word line-for-line, the ~1-page TOC over-production
  is NOT in the primitives — it is real-entry-specific: candidates are extra
  wraps on the 6 real titles >55 chars (ALL-CAPS TOC1 headings measure wider) or
  the cached `PAGEREF` field's page-number-column width vs a literal number
  (the probe used literals). Pinning it needs the REAL fixture exported through
  Word, which is currently blocked NOT by a locked session but by the fixture
  FAILING validation (numId=0 undefined + `<w:shadow>` out of order in `<w:rPr>`)
  — opening it risks a repair dialog that taints refs. Sanitize the fixture
  (fix numId/rPr order) first, then export and diff the real TOC's wrap columns.

### allowOverlap="0" slides an anchored shape clear of earlier overlapping floats
`wp:anchor @allowOverlap="0"` means Word shifts the shape so it does NOT overlap
any earlier-placed (lower z-order) float, rather than stacking on top
(staging-anchors2: the locked, no-overlap z=30 box is authored at page-x 144pt but
Word slides it right to ink-x 292pt — its left edge lands at the z=20 box's right
edge ~273.6pt). Parsed onto `ShapeTextbox.allowOverlap`; `emitAnchors`
(`layout/engine.ts`) tracks placed rects in z-order and, for an `allowOverlap===
false` shape, shifts `ox` right past any vertically-overlapping earlier rect.
Fixed the z=30 box position exactly. anchors2 p1 residual is now the
relH="character"/relV="line" anchor (Word resolves it to the inline character's
x/y; we still treat character-relative as column-relative), an unimplemented
anchor coordinate system.

### A picture's `a:ln` outline paints just OUTSIDE the image, even widthless
`<pic:spPr><a:ln>` with a solid fill (no `noFill`) is a picture border Word draws
as a line straddling the image edge — a widthless `a:ln` (no `@w`) still paints at
a ~0.8pt hairline (hamburg "Marketing" figure: black `tx1` outline, rule bbox
sits half its width beyond the image box). The pic parse ignored it. Fix: parse
`a:ln` (guard `!noFill` + a resolved fill color; `tx1`→`dk1` via theme aliasing)
into `DrawingImage.border`/`ImageContent.border`, thread it through the inline
atom (`layout/inline.ts` — it must be copied at BOTH the atom builder AND the span
repack, plus the `ImageAtom` type) into `ImageItem`, and render it as a CSS
`outline` (draws outside the box without shifting the image or its layout). Fixed
hamburg p10 5.1→1.1.

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

### Trailing spaces hang past the line end (never start the next line)
Word absorbs every space at a line boundary into the END of the wrapping
line: the spaces hang invisibly past the margin with real advances, they
never affect alignment/justification or line metrics, and the next line
never begins with a space. An earlier fix carried "extra" wrap-boundary
spaces to the next line's start — that pushed the next word right and made
the caret leap backwards to the lower-left when typing a space at a wrap
boundary (the user-visible "backwards space" bug). A trailing space at a
paragraph end likewise must emit a span, or the caret loses its binding
after the re-render and vanishes. Implementation: `flush()` in
`layout/inline.ts` detaches trailing space spans before
`finishLine`/`applyAlignment` and re-attaches them sequentially past the
line's visual end (bidi paragraphs keep the old drop). Pixel-parity neutral
by construction: the spans paint no ink and alignment math never sees them.

### Caret affinity at wrap boundaries (Caret.bias / SelPoint.bias)
A caret offset at a span boundary is ambiguous when the two spans sit on
different lines (soft wrap). Word resolves it by HOW the caret got there:
typing/Backspace/ArrowRight/End/clicking past a span's right edge show the
caret at the end of the UPPER line; clicking at a lower line's start shows
it there. `positionCaret`'s binding pick honors `bias: "end"` (prefer the
span ENDING at the offset). Also: Range rects inside whitespace-only spans
lie (the DOM collapses trailing-space widths to zero), so the caret x for
any whitespace-only span comes from the layout item's own geometry.

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
- **A first line that physically won't fit moves the WHOLE paragraph — even
  with widowControl OFF**: pushing a paragraph down when its first line doesn't
  clear the body bottom is a PHYSICAL fit, not the aesthetic orphan rule, so it
  applies regardless of `w:widowControl="0"`. `planBreaks` couldn't record a
  break at the very first line (its `li > segStart` guard blocks `0 > 0`), so it
  planned the break AFTER line 0; the emit loop's own overflow test then moved
  line 0 to a fresh page, but the STALE post-line-0 break still fired there,
  orphaning a lone first line on an otherwise blank page and spilling the rest
  onto the next. Word's default widowControl=on masked this (the orphan cascade
  set the break to 0), but `Default`-styled bodies with widowControl=0
  (nccih-protocol's SOM/eligibility notes) tripped it: 3 spurious near-blank
  pages (26→23), mean 64.3%→24.3%. Fix: `planBreaks` detects the
  first-line-overflow-on-a-partial-page case and breaks at 0 (whole paragraph to
  the next column/page) independent of widowControl, so plan and emit agree.
  Guarded off during column balancing (that path is authoritative there).
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
- **w:beforeAutospacing / afterAutospacing = a FIXED 14pt margin, ignoring
  the literal before/after AND the font size** (HTML/web-pasted content,
  `NormalWeb`, FEMP RFP bracketed guidance blocks): Word discards the
  `w:before`/`w:after` and inserts a constant 14pt (CSS px `14 * 96/72`)
  above/below the paragraph, NOT the paragraph's own line height. Measured
  across wild-doerfp's 10.5pt guidance blocks (three baseline boundaries:
  afterAuto = 14.03 / 13.75 / 14.00pt) and wild-athabasca's NormalWeb title
  page (27.8pt gaps = 13.8pt line + 14pt auto). The earlier `naturalHeight`
  rule only worked because athabasca's font gave a ~13.8pt line ≈ 14pt; for
  sub-12pt paragraphs it undershot ~2.3px per boundary (doerfp section pages
  stacked TWO such boundaries — title→guidance + guidance→body — into a
  ~6.6px whole-body shift → p31 15.1%→5.2%, p32 14.0%→3.5%). The fixed value
  self-satisfies the "double spacing (line=480) must NOT inflate the auto
  gap" rule since it ignores the line-spacing multiple entirely. Floored at
  `naturalHeight` so a rare large-font autospacing paragraph never gets less
  than one line (`AUTO_PARA_SPACING_PX` in engine.ts). **A trailing EMPTY
  autospacing paragraph's
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

### The Times-family "cumulative drift" in the wild2 math fixtures is NOT a font-metric error — it is localized math/equation/image construct height (2026-07)
The three math-heavy wild2 fixtures were suspected of a uniform Times/TNR
per-line vertical drift (~+0.2pt/line) compounding to a full page of pagination
desync (web 18 vs Word 17 on `wild2-math-omml-dense`; 7 vs Word 8 on
`wild2-math-eq-as-images`; `wild2-sci-chem-omml` 13 = 13 but pages 3-13 badly
misaligned). Direct measurement of Word's own reference PDFs **falsified** the
uniform-metric hypothesis:
- **Pure body-text TNR pitch matches ours to ~0.015pt/line.** Char-baseline
  clustering of Word's PDFs gives a dominant 12pt single-spaced pitch of
  **13.75pt** (raw ≈ 13.78-13.80) vs our raw `1.149902×12 = 13.799`; the
  chem doc's double-spaced (`line=480 auto`) 12pt pitch is **~27.6pt**
  (= 2×13.80). On a clean 72-line double-spaced body stretch the ourY−wordY
  drift moves only −1.1pt (−0.015pt/line). The larger 14.0/14.5pt deltas in the
  histograms are inline-math/subscript lines whose line box legitimately grows,
  not the font pitch. `WORD_FONT_METRICS['times new roman']` is correct — do
  NOT tweak the hhea value (it would regress the verified-correct body pitch;
  msa/chronology are Arial so they wouldn't catch it).
- **The drift is localized and per-construct, in OPPOSITE directions per doc.**
  Matching identical (sanitized) text lines between our render and Word's PDF
  and plotting cumulative ourY−wordY: `dense` holds a flat ~−28pt offset through
  the text pages then takes discrete positive JUMPS at display equations/figures
  (+99, +254 at single blocks) totalling +790pt ≈ one extra page — OUR tall
  constructs over-size Word's. `eq-as-images` plunges NEGATIVE at each equation-
  image block (−162, −259, −356 …, ending −394pt ≈ half a page short) — OUR
  equation images UNDER-size Word's, so we pack one page fewer (7 vs 8). `chem`
  nets ~0 (13 = 13). A single font metric cannot move pagination in two
  directions at once; the real fixes live in the math-cluster / display-equation
  / drawing-image height code and need Word-probe calibration per construct.
- **Real latent bug fixed on the way: bare "Times" was un-mapped.** The dense
  fixture's `docDefaults` ascii family is bare `"Times"` (not "Times New
  Roman"); ~27% of its text runs (957 of 3531) inherit it. `"times"` had no
  entry in `WORD_FONT_METRICS` **or** `METRIC_SUBSTITUTES`, so those runs fell
  through to canvas `fontBoundingBox` (integer-rounded — the same failure mode
  as the Arial 16-vs-16.87 note above) for BOTH advances and line height. Word
  resolves "Times"→"Times New Roman"; adding `times` to both maps (→ TNR metric
  + TNR substitute) flattened the dense front-matter drift from a growing
  −42pt to a constant −28pt (the residual is the localized-construct offset).
  Page counts and doc means are unchanged because they are dominated by the
  construct-height desync, but the geometry is materially more correct and the
  mapping is the right resolution regardless.
- **Method note:** measured entirely from the EXISTING Word reference PDFs —
  a fresh Times probe (`scripts/make-times-probe.py` / `read-times-probe.py`,
  covering TNR + bare "Times" at 10/10.5/11/12pt × single/double/1.08/atLeast)
  was generated and validated but could NOT be exported through Word: past
  midnight the session was screen-locked, so every `open`/save-as AppleEvent
  timed out with −1712 and no window ever appeared (the documented
  "exports need an unlocked session, full stop" failure). The probe is ready to
  export and regress once the construct-height work resumes on an unlocked box.

### The dense-math pagination desync was a `w:position` raise inflating a figure line beside its inline image (2026-07)
Following the construct-height diagnosis above, the dominant erring rule in
`wild2-math-omml-dense` (web 18 vs Word 17) was found and fixed. The doc's
figures are laid out as an INLINE picture followed by a same-line label run
(`h_0-05.jpg` + "V1"/"R1"/"J1"), where the label carries `w:position w:val="320"`
(**160pt = 213px raise**) at `sz 32` (16pt). Word's line box for that paragraph
is the **image height** (~248px): the raised 16pt label's top (raise +
glyphAscent ≈ 228px) stays INSIDE the 248px image extent, so it adds nothing.
`finishLine` (`layout/inline.ts`) instead added the FULL raise ON TOP of the
line's natural height — and that natural was the image height, not a text line —
so each figure block rendered **465px instead of 252px** (+160pt each × 3 = the
prior "+254pt at figures" jump), pushing one extra page.
- **The `w:position` line-box rule is `max(objectHeight, textAscent+raise)`, not
  `objectHeight + raise`.** The old model added the shift as pure extra leading
  above the natural line — correct for a text-only line (the charstyles probe:
  +6pt raise = +6pt pitch, because there natural IS the text line), but wrong
  when a taller co-line object (inline image/drawing) already covers the raised
  glyph. Fix: track `maxNonObjAscent/Descent` (text + inline-math only) and
  resolve the raise as the amount the raised/lowered TEXT protrudes past the
  line's overall ascent/descent — `raiseAsc = max(0, maxNonObjAscent + raise −
  maxAscent)`. For a text-only line `maxAscent === maxNonObjAscent`, so it is
  still the exact full shift (0 mismatches on the 6 line-break fixtures, all
  math/display parity fixtures unchanged: parity-math 5.03, parity-math2 4.16,
  parity2-equations 8.50, benchmark 2.83, gatech 1.99, hamburg 1.46). For the
  figure line the image ascent wins and the raise contributes 0.
- **Result:** all three math wild2 page counts now MATCH Word — dense 18→**17**,
  chem 13=**13**, eq-as-images already 8=**8** (a prior fix; the memory's "7"
  was stale). Dense's figure pages went from 88/90/99% structural severity to
  **9/4/17%** and its doc mean dropped **69→55**. The residual dense/chem/
  eq-as-images means (55/69/77) are NOT construct height — the display-equation
  block heights measure within Word's own tall-equation line gaps (ours ~21–31pt
  vs Word ~23–35pt for the ∫/∑ blocks on p8) — they are the STIX-Two-Math ↔
  Cambria Math glyph substitution shifting every math glyph horizontally, which
  the structural comparator scores heavily. That is a font-rendering gap, not a
  height rule, and is out of scope for pagination parity.
### wild2-legal-nih-contract severity is diffuse cumulative TABLE-ROW drift, not one wrong rule (2026-07)
The 419-page NIH contract (we render 417, baseline mean ~48% over the first 286
pages before a Playwright screenshot timeout, 62% of pages ≥50%) was suspected of
a specific long-table row-pagination bug. Direct measurement from the reference
PDF **falsified** the single-bug hypothesis and pinned the mechanism as a broad,
sub-pixel accumulation:
- **The severity mass IS a cumulative page offset.** Aligning our body-line
  stream to Word's (difflib on sanitized text, tracking each matched line's page
  number) shows our pages progressively run AHEAD of Word — content on the same
  index page drifts from 0 → −1 (by ~p29) → −2 (by ~p238) → −3 (by ~p290), ending
  −2 (417 vs 419). Because parity-compare scores SAME-INDEX pages, a sustained
  −2/−3 offset makes essentially every back-half page ~100% mismatched. The
  baseline bands confirm it: p1–240 mean ~43%, **p241–280 70.7%, p281–320 81%** —
  the severity tracks the drift magnitude, not any local content, so "our pages
  265–274 all 100%" is the accumulated offset surfacing, not a local table bug.
- **The doc is 728 content-sized tables / 2095 rows, ZERO `w:trHeight`.** So every
  row height = cell content + cell margins (this table style: top/bottom = 0,
  L/R = 108tw) + `rowBorderShare` (sz6 = 0.75pt all sides). 56% of paragraphs live
  in cells; the body is a lattice of one-block tables carrying "RUH MUKOR"
  (`****(USE MODEL…)`) bracketed-guidance rows, bold-red number paragraphs, and
  lettered clause items.
- **Every construct, measured individually, matches Word to sub-pixel.** Clean
  single-row schedule tables (p15) match Word's row rules within ~0.05pt/row. The
  12pt Calibri body pitch: Word's dominant 14.75/14.50 gap alternation averages
  **14.652pt = ours (14.648)** — the 14.726 overall mean is inflated by
  legitimately taller content lines (fill-in-blank underlines, tab leaders,
  sub/superscript), NOT a pitch error (do NOT tweak the Calibri metric; it would
  regress benchmark/sample). Multi-line guidance rows resolve `before=15tw`/
  `after=25tw` correctly (engine probe: 3-line frame 61.26px ⇒ 46.7pt row vs
  Word 46.75). The only single-line guidance rows that render short (15.00 vs
  Word 15.50) are the ~10 rare spacing-less variants, not the 920 normal ones.
- **The residual is a diffuse ~0.1–0.3pt/row deficit** spread over content-heavy
  lines Word sizes slightly taller and over body-paragraph→table boundaries (the
  number-para→guidance-table gap is Word 16.25pt vs ours 15.36 = −0.85, but it
  decomposes into two LARGER opposing sub-pixel effects — Word's table top border
  sits ~2.2pt higher AND its cell text ~0.5pt lower — so it is not a clean
  row-height error). ~3–4pt/page × 419 pages ≈ 2–3 lines lost per ~4 pages ⇒ the
  ~2-page shortfall and the −3-page mid-document drift. No single row rule
  (`rowHeightFromTrHeight` is unused here, `rowBorderShare`/`cellMarginsOf` are
  correct for the clean rows, `ROW_OVERHANG_TOL` reclaims our own drift) is wrong.
- **Method / plateau.** Measured entirely from `parity/wild2-legal-nih-contract-
  word.pdf` (pdfminer row-rule + baseline geometry) vs DOM probes on :5315 and a
  temporary `layoutRow` instrumentation (removed). Pinning and safely correcting
  the per-construct ~0.1–0.3pt deltas needs a Word-export probe sweeping this
  doc's cell-margin/border/fill-underline line heights — which the overnight
  screen-lock (−1712, no window) forbids. No engine change is committed: any
  row-height nudge that "fixes" this doc's page count without a Word reference
  would be fitted to our own measurements and risks the five calibrated table
  fixtures (staging-longtable, parity-rowsplit, parity-tables, parity2-
  nestedtables, and the row rules' benchmark/sample anchors). This is a genuine
  plateau pending an unlocked Word box, same family as the wild2-math construct-
  height backlog above.

### NIH row-height probe: the deficit is NOT per-row — it is a discrete ~0.79pt per PARAGRAPH→TABLE boundary (2026-07, unlocked Word)
The plateau above was finally probed on an unlocked box.
`scripts/make-nih-rowheight-probe.py` reproduces the fixture's exact guidance-
table style (docDefault Calibri, Normal sz=24; number-paragraph = keepNext +
`before=100` bold-red; guidance table = tblW auto, tblInd 500, ALL borders
single **sz6** (0.75pt) space0, shd F3F3F3, single gridCol 9700, cantSplit rows,
cells inherit TableNormal cellMar top/bottom=0 L/R=108, cell paras
`before=15`/`after=25`) in five blocks: P 50 single-line rows, Q 20 three-para
rows, R 40 `[number-para][1-line table]` units, S 20 `[number-para][3-line
table]` units, T 40 underline-fill rows. Exported to
`parity/probe-nih-rowheight-word.pdf`; Word pitches read with
`read-nih-rowheight-probe.py`, ours with `read-nih-ours.mjs` (DOM, :5317, px×0.75).
Both regress baseline/top-y vs index so the constant baseline-vs-lineboxtop
offset cancels and the pitch deltas are reference-independent.

| construct | Word | Ours | Δ (Word−Ours) |
|---|---|---|---|
| P single-line row pitch | 17.403 | 17.398 | +0.005 |
| Q 3-para row pitch | 49.209 | 49.195 | +0.014 |
| T underline-fill row pitch | 17.402 | 17.398 | +0.003 |
| Q intra-cell line pitch | 15.906 | ~15.875 | +0.03 |
| **R unit pitch (numpara→numpara)** | 37.803 | 37.047 | **+0.757** |
| **S unit pitch (tall table)** | 69.612 | 68.843 | **+0.769** |
| **R para→table boundary gap** | 16.256 | 15.466 | **+0.790** |

- **Intra-table row pitch is CORRECT** (Δ ≤ 0.014pt across single-line,
  3-paragraph, and underline-fill rows). Underline-fill runs add **zero** height
  (T === P). The "diffuse ~0.1–0.3pt/row" hypothesis in the entry above is
  therefore **wrong** — there is no per-row deficit.
- **The entire drift is a discrete ~0.79pt shortfall at each PARAGRAPH→TABLE top
  boundary** (numpara baseline → first cell line). Over the fixture's 728
  one-block tables that is ~575pt ≈ 1.4 pages — the dominant part of the
  observed −2p (417 vs 419) drift, direction confirmed (we pack tables too
  tight ⇒ run short). The table→following-paragraph (bottom) boundary is ~right
  (R/S unit deficit ≈ the top-boundary deficit alone).
- **Decomposition (two opposing sub-pixel effects, matching the prior note's
  hunch, now measured):** From the PDF, Word's table top border sits **1.0pt
  below** the preceding numpara baseline (705.47→704.47) — i.e. it OVERLAPS the
  paragraph's descent. Our engine drops the border to the numpara's full
  line-box bottom (~3.2pt below baseline), so **our border is ~2.2pt too LOW**.
  Independently, Word seats the first cell line **15.25pt below the top border**
  (identical in block P which has no preceding paragraph, so it is a pure
  cell-internal rule), whereas ours seats it ~2.5pt higher, so **our first-cell
  line is ~2.5pt too HIGH relative to the border**. The two nearly cancel, net
  +0.79pt in Word.
- **No fix shipped — deliberately.** A correct fix must change TWO things at
  once (table-to-preceding-paragraph attachment must overlap the paragraph
  descent; the first line inside a cell must reserve ~a full line's top leading
  below the top border, not just the glyph ascent) and net exactly +0.79pt.
  Getting only one right over/under-corrects, and both touch every table, so it
  must be gated on staging-longtable/parity-rowsplit/parity-tables/
  parity2-nestedtables + benchmark/sample before landing. A bare +0.79pt nudge
  on para→table boundaries would be a fitted constant (the anti-fitting rule
  forbids it). The probe + both readers are committed so the two-part rule can
  be built and validated safely; this is now a *characterised* task, not a
  blind plateau.

### NIH row-height probe CLOSED: the residual was WIDTHS and KEEPS, not row heights (2026-07)
With the para→table boundary rule already landed, probe-nih-rowheight still
scored mean 3.10 and the wild2 doc kept a −2p drift. Rule-level measurement of
the probe PDF (pdfminer rules vs our DOM rules) pinned five independent causes;
fixing them took the probe to **0.00 on all 12 pages**, the wild2 page count to
**419 = Word**, and the first 50 pages to ≤2.5% (TOC pages 3-8: 0.1-0.3%):
- **A tblW=auto table whose trusted grid overruns its slot is CLAMPED, not
  hung.** The probe's guidance table (gridCol+tcW 9700tw, tblInd 500tw, 9360tw
  column) renders 443pt wide in Word (= column − indent; rules span x 97.4 →
  539.6), not the authored 485pt. Only an EXPLICIT dxa tblW hangs (gatech's
  9129tw table — that rule stands, now conditioned on `width !== undefined`).
  In the real fixture the margins are 900tw so the 9700 grid FITS — the probe
  had mis-reproduced the section and this clamp never fires there.
- **CRACKED: Word's shrink rule for over-wide tables — col = tcW − (tcW −
  minContent)·k, k = (ΣtcW − T)/Σ(tcW − min).** For a pct table whose per-cell
  tcW total exceeds the pct target, Word IGNORES the cached tblGrid and shrinks
  every column from its tcW preference toward its min-content width
  proportionally to the slack. Verified against rendered rules to ≤0.2pt on
  wild2's p16 5-col financial table (predicted [151.0,78.4,64.3,66.2,89.0] vs
  measured [150.83,78.52,64.28,66.02,89.03]; the cached grid is 5.3pt off) and
  p17's 6-col (model ≤0.2pt, grid 10pt off). Paragraph left-INDENTS count
  toward min-content (p19's ind=720 headers). Fresh cached grids equal the
  model's output — stale ones (cells edited after the last relayout) do not,
  so the model, not the grid, is authoritative (`shrinkToTargetWidth`).
  The wrap-critical consequence: our col0 was 5.3pt wide, so "[Koja Mugevu
  and Nuhiha(n)]" fit on one line where Word wraps it — one lost line per
  financial table, cascading into cantSplit/keepNext flips at page bottoms.
- **A space run whose next word starts with an NBSP is NOT a break
  opportunity.** Fill-in blanks ("of $ [12×nbsp] (lohirol)") wrap as ONE glued
  unit in Word — it moves "$" down with the underlined NBSP run instead of
  ending the line at "…of $" (SpaceAtom.noBreak; head/tail walks glue through).
- **keepNext binds a paragraph to a following TABLE's lead block** (top border
  half + leading tblHeader rows + first data row — `tableLeadHeight`), and
  **tblHeader rows never sit alone at a column bottom** (header + first data
  row move together). But a LONG keepNext paragraph (4+ lines) does NOT move
  whole: Word splits it like any paragraph and only its final line (+ widow
  companion) binds forward (p34/35: [3×w:br + "58"] leaves two break lines at
  the page bottom). Implemented as a planBreaks tail reservation on the final
  line (`keepNextTail`), with long chain members terminating the chain walk.
- **cantSplit yields when the row is taller than one page** — Word moves it to
  a fresh page first, THEN splits it mid-row (p115/116 giant guidance row; we
  used to let it overflow past the page edge).
- **A right/decimal tab whose aligned text cannot reach its stop wraps to a
  fresh line and re-evaluates** (full leader + number at the stop), while text
  ending within ~0.25pt of (stop − numberWidth) right-aligns normally — the
  TOC's "…CUQIKAPUBAK126" pack vs "…KIPULAMURA" + "……… 220" wrap differ by
  advance-exact +0.11pt vs +1.31pt past target (tolerance 0.75pt). The aligned
  run is POSITIONED, never re-wrapped (no 2px minimum tab width, overhang past
  the content edge allowed) — a forced 2px push made stop-adjacent numbers
  wrap bare to the left margin and desynchronized every TOC page after p3.
- Residual: pages ~116-260 keep a local ~25-35% band from borderline
  cantSplit/keepNext flips in the URL-dense schedule region (Word breaks long
  URLs at path separators; we char-wrap — see the hyphenBreaks note). Both
  ends of the doc re-sync (p1-50 ≤2.5%, p270-419 ≈2-6%), PAGEREFs match, and
  the suite gates are unchanged (doerfp 0.81, athabasca 0.89, longtable 1.68,
  parity-tables/benchmark/sample/charstyles 0.00, nestedtables 0.17,
  compare-linebreaks 6 fixtures unchanged).

### NIH 116-260 band CRACKED: Word's URL-break rule measured from the PDF, plus two numbering root causes (2026-07)
The reference PDF itself was the probe: every mid-token line break on pp116-260
was extracted (pdfminer) and matched against the docx's unspaced tokens. 22
genuine break decisions survived validation (12 apparent "break after '.'/':'"
cases were w:br artifacts — the sanitized docx splits `word.</w:t><w:br/>
<w:t>Word` and naive w:t concatenation invents unspaced tokens; ALWAYS check
the raw XML for w:br before trusting a corpus pair). The measured rule:
- **The ONLY soft break inside an unspaced token is after a hyphen with
  alphanumerics on BOTH sides — digits included.** ".../GUF-JE-" | "04-332"
  (letter-digit) and ".../h44-" | "40.aki" (digit-digit) break; the old
  letters-only hyphenBreaks missed both. Leading minus ("-4") still excluded.
- **'/', '_', '.', ':', '?', '=', '&' are NEVER break opportunities.** Word
  char-wraps PAST them ("…Corinazib/Ha" | "rujipaguduh.loh",
  "…BOB_HUG_Kudifup" | "a_Sucumo.idi"); the p157/p164 breaks that LOOK like
  break-before-'&'/'=' are exact-edge char wraps (next char overflows by
  width). This is why the earlier eager '/'+'_' experiment scored worse.
- **Emergency break happens IN PLACE at the exact overflow character.** When
  no opportunity exists on the line (the glued unit — NBSP glue included —
  reaches back to line start), Word fills to the edge and continues on the
  next line at char granularity. It does NOT flush the glued head ("at:" +
  NBSP) to its own line first (p154: "at:  wamuv://…BOB_HUG_Kudifup" ends
  75pt short of the page edge because the para carries ind right=1440 — the
  wrap is at ITS line edge). Implemented as hardWrapFrag + a hi===minSpans
  early exit in the head-walk (never beside a float).
- 22/22 corpus breaks match the implemented rule (15 letter-hyphen already
  handled, 2 digit-hyphen new, 5 emergency char wraps new).
Two NON-URL root causes drove most of the band's page drift:
- **A keepNext chain-walk measurement consumed once-only startOverride
  restarts.** The walk snapshots counters but (before this fix) not
  seenNumIds, so the restart fired during measurement, the counter rolled
  back, and the real placement never restarted: numId 340 rendered hh/ii/jj/
  kk where Word shows a/b/c/d, shifting label widths and wraps. All counter
  snapshot sites now roll back seenNumIds too.
- **w:lvlJc="right" was ignored.** Word right-aligns the label at the number
  position (ind.left − hanging): the label's RIGHT edge sits there and grows
  leftward ("i." → "viii." all end at the same x), so the suffix-tab text
  NEVER moves off ind.left. Our left-aligned labels pushed text to the next
  default stop for wide labels, wrapping "the kohi"-class words one line
  early across the whole section (p177/178 cascade).
Band after: spot set 30/50/80/120/135/150/165/177/180/185/193/200/210/225/230
all ≤3.8% (was 25-36 at 165/180/200). Remaining pockets, characterised and
NOT URL-related: p189-192 (one borderline keepNext flip at the p188/189
boundary — the heading has ~53pt of room in Word vs our ~56pt need; hinges on
the deferred para↔table boundary sub-point metrics) and p244-265 (the FUZ
clause-matrix table's title column renders ~1 word wider than Word's — table
width model, tracked separately). Gates: unit tests 184 green,
benchmark/sample/charstyles 0.00, ca-agreement 0.16 mean, longtable and
doerfp byte-identical to the pre-change baseline on every page,
compare-linebreaks sample/chronology/msa 0/0/0.

### NIH residual mass CRACKED: footer-frame phantom line, pct-table re-autofit, multi-space/NBSP wrap glue, unsplittable keep terminators (2026-07)
Four independent rules, all measured from the wild2-legal-nih-contract Word
PDF, took the fixture from mean 9.89 to ~0.3 (see run history):
- **Widthless PAGE-frame footer: the empty ptab follower SHARES the frame band,
  and the frame adds a HEIGHT-ONLY phantom line once its text is wider than
  its glyph box.** Word footer tops across all 419 pages: pageBottom −
  footerDist − 3 lines on pages 1-9 (one digit), − 4 lines from page 10 on
  (two digits) — while the PAINTED stack (number, admin line exactly one line
  below) never changes. The phantom moves the footer anchor AND the body
  bottom (bodyBottom = pageH − footerDist − footerH), which was both the TOC
  drift (p2-6: one missing body line per page ≈ 3 entries behind by p5) and
  the near-blank landscape pages' 90% scores. Our glyph box is win-metrics
  (~1.22em) so the width test is 0.7×boxH: one digit 8.1px stays under, two
  digits 16.2px clear it. Do NOT test `boxH − 3`: it silently excludes
  2-digit pages and hands every p10-99 page an extra body line (5-page
  pagination landslide by p285).
- **A pct-width table with a trusted grid is re-autofit: columns rise to
  min-content, funded by slack columns (col = raised − (raised−min)·k).**
  The FUZ clause matrix (tblW 4800 pct, grid [1394,1193,7435]tw) renders
  [76.02, 59.28, 365.82]pt in Word — col1 = the NBSP-glued " FETOWO GO. "
  header min — not the grid's [69.7, 59.65, 371.75]. Use Word-exact mins
  (subtract columnMinPref's +2px fudge) or col2's NBSP-glued "Wej 7426"
  (59.25pt) gets a raise Word doesn't do. The raised col1 also makes the
  repeated header row THREE lines (leading space + NBSP-glued chunk no longer
  fit one line — see next item), which Word shows too.
- **Word treats any whitespace cluster containing 2+ spaces (or an NBSP
  touching a space) as NON-BREAKING: the flanking words wrap as one unit.**
  Measured twice: 'Hunogigu."\xa0 Durirone' moves to the next line as a
  108.6pt unit though 'Hunogigu."' alone fits the 99pt remainder (p106), and
  'nuqagajote␣␣' + 80 underlined fill-in spaces wraps before "nuqagajote" as
  a 279pt unit into a 275.6pt remainder (p383). Implemented as noBreak glue
  on space spans adjacent to spaces, plus the mirror of the existing
  "next word starts with NBSP" pass for words ENDING in NBSP. Also: a space
  directly after an explicit <w:br/> is REAL line-initial content (it
  consumes width; NIH header " FUZ <br> FETOWO GO. " wraps into a space-only
  middle line) — cleared p143-146. compare-linebreaks sample/chronology/msa
  stay 0/0/0.
- **A 2-3 line keepNext TERMINATOR is unsplittable under widow control
  (2+1 widows, 1+2 orphans), so the chain must reserve ALL of it.** NIH
  p416/417: '537' (keepNext) + Heading4 + 3-line URL paragraph move to p417
  as one 79pt block leaving 90pt unused; reserving only first+widow lines
  strands the headings at the page bottom (p416-418 at 47/30%).
Remaining known-off: p343/344 flip rides ~1.4pt of accumulated sub-point
list spacing (the documented deferred para spacing class; +0.25-0.3pt per
bullet item through p342's Qizunuroqufa list), and the footer admin line
paints 5pt right of Word (our trailing-tab + jc=right resolution vs Word's
333.40 — ~9.8% floor on the six near-blank pages, invisible elsewhere).

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
  not the 624px content width. Open calibration problem, parked. Re-measured
  2026-07: Word's 601px is NOT plain content-fit either — the two content
  columns (Key 24.5pt, Status 40.6pt, both content+slack) match if our autofit
  `pad` shrinks ~2.1pt/column, but the flexible description column then makes the
  table 601px, between our content-min (~445px) and full width (624px). The 601
  derivation stays uncracked without a Word-export probe sweeping content ×
  margins × the flexible-column share.
- **CRACKED (2026-07, probe-autofit): Word autofit is simply
  PROPORTIONAL-TO-PREFERRED-WIDTH; there is no per-column "slack".** A 10-block
  sweep (`scripts/make-autofit-probe.py`: 2–4 cols, single non-breaking tokens
  of known width, cell margins 0/108/300tw, borders on/off, tblW pct=5000 vs
  auto) exported to `parity/probe-autofit-word.pdf`; column boundaries read from
  the vertical rules (`read-autofit-probe.py`) vs our DOM (`read-autofit-ours.mjs`).
  Modelling `col_i = prefᵢ · T / Σpref`, with `prefᵢ = (content advance) +
  (cell L+R margins) − ~2.4pt` (the ~2.4pt is the glyph SIDE-BEARING my
  pdfminer glyph-extent over-counts vs Word's advance width), reproduces EVERY
  block to <0.7pt — e.g. 4-col F predicted 41.4/84.4/142.1/199.8 vs Word
  41.2/84.3/142.3/199.9; custom-300tw-margin G 107.4/360.2 vs 107.5/360.2;
  zero-margin E 46.0/116.7/305.0 vs 46.7/116.8/304.2. Equal-content tables (A,J)
  always split T equally regardless of content size. So the parity-tables
  "non-constant slack (10.4 vs 15.9px)" was an ARTIFACT of diffing against raw
  content under a proportional (not additive) distribution plus side-bearing
  measurement error — not a real per-column rule.
- **Our engine already distributes proportional-to-pref** (`resolveGridWidths`:
  `prefW[i]*want/sumPref`), so the DISTRIBUTION is correct. The residual is that
  our `columnMinPref` `pad = margins.L + margins.R + 2` adds a FLAT `+2px` and
  our content measure (Carlito advance) INCLUDES the side bearings Word drops, so
  our `prefᵢ` over-weights narrow columns; on the probe our narrow columns run
  ~1–5px wide and the wide column ~equally narrow (worst: zero-margin block C
  67.5 vs Word 63.0; E 49.5 vs 46.7). Direction is uniform (narrow-too-wide).
  A fix = shrink narrow-column pref bias (drop the `+2`, and/or subtract a
  side-bearing estimate from measured content) — but it perturbs EVERY autofit
  column width and thus line wrapping, so it is gated on the full line-break
  suite (compare-linebreaks 6×0) + parity-tables/staging-grid4/staging-tblextreme
  + benchmark/sample and was NOT shipped blind here. parity-tables' remaining gap
  is separately the invalid-`w:w="100%"`-pct (Word ignores → content-fit ~601px);
  valid `type="pct"` fills 100% correctly in both engines (probe totals both 467.7).

## Sparse-page metric floor (near-blank pages score high on the structural NCC)
On pages with very little ink (parity-dividers 8.4%, pleading p4 9.0%, gatech p10
5.4%), the structural severity is dominated by Chrome's ~1.1-1.2x heavier text
rasterization and sub-pixel rule antialiasing, NOT by geometry: every element
(line numbers, rules, headings, body) verified to match Word within ~0.5px, yet
the tile-weighted NCC has little other ink to dilute the weight residual (raw
mismatch stays ~0.6-1.3%). These are a plateau, not a bug — moving them needs a
global `-webkit-text-stroke`/baseline recalibration that the tuned
benchmark/sample/pickett fixtures forbid (and the per-page weight ratio VARIES:
dividers 0.96x, pleading p4 0.98x, gatech p10 1.18x, so no single global constant
fixes them). gatech p10 also carries a genuine REF/SEQ field-value divergence
("Bavoqe 0" cached in Word vs our live-evaluated "1") — a field-resolution gap,
distinct from the weight floor.
- **The text-ink weight floor is NOT tunable via stroke calibration** (swept
  2026-07-09 over 48 diverse pages: regular antialiased+0.15px/bold subpixel
  = 133.6 total severity; 0.05px = 134.9; no stroke = 135.8; bold-as-stroke
  0.3px = 134.3). The current calibration is the optimum; the residual ~2-4%
  "clean floor" on sparse pages is Chrome-glyph-rasterizer vs Word-PDF-
  rasterizer AA physics at identical glyph positions (both sides already
  compare at 2x scale). Do not chase it with paint tweaks; it caps the
  achievable suite mean around ~2.2-2.5% under the current severity metric.


## Legacy CJK docGrid layout: grid-snapped object/text lines, EA space fitting (wild2-math-eq-as-images, 2026-07)

The 8-page Chinese physics paper (compat mode 12, docGrid type=lines
linePitch=312 = 15.6pt, SimSun body 10.5pt, HTML-Preformatted paragraphs at
before/after 156 + line 348 atLeast, every display equation an inline VML/OLE
image on a `w:position`-lowered run) went mean 40.3% -> **0.38%** (worst page
74.4% -> 0.67%) from rules measured directly in its Word PDF (paragraph
shading rects + baselines + image bboxes agree to ~0.1pt):

- **Grid OBJECT lines snap to whole pitches, centered.** A line whose inline
  object extent (ascent+descent split by the run's w:position) exceeds the
  GRID PITCH - not the paragraph's spacing height - lays
  `ceil(extent/pitch) x pitch` with the extent centered: img 31pt -> 31.2
  (2 pitches), 36 -> 46.8, 48-57 -> 62.4. eq48's image top sits exactly
  (62.4-57)/2 = 2.7pt below the line top (shading rect 643.44, image 640.75).
  A 15pt image lowered 3pt (extent exactly 15 <= pitch) keeps the plain
  atLeast line; the same image at position 0 (extent 16.5 via text descent)
  snaps to 2 pitches. `w:position` moves the image across the baseline but
  the snapped height is unchanged while the image dominates both sides.
- **Legacy (compat < 15) grid TEXT lines snap too** when the font line
  (with gap) exceeds the pitch: the sz28 headings (SimSun 14pt = 15.97pt
  line) lay 2 pitches = 31.2 under atLeast 348 AND under plain auto
  multipliers (CSO- 1.25, Normal 1.0) - the snap REPLACES the multiplier.
  staging-eastasian (compat 15, same grid type) does NOT snap its oversized
  faces, so the text snap is gated to compatibilityMode < 15 (and never
  fires for oversized East Asian faces, which keep multiplier x natural).
- **VML pict extents round to whole points** (31.45->31, 49.65->50,
  120.75->121, 290.75->291; both axes) - the PDF draws every equation raster
  on integer pt, and the rounding decides 2 vs 3 grid pitches for the
  31.45pt images.
- **A word binds to a directly attached VML/OLE object** (no space): Word
  wraps "as:<eq>" as one unit even though "as:" fits (line1 ends at x=326 of
  505). DrawingML pictures do NOT glue (chem p3 keeps its "[06]" marker on
  the line and wraps the chart alone).
- **East Asian line fitting compresses inter-word spaces** (docGrid docs,
  any alignment): left-aligned SimSun lines draw every space at
  5.00/4.25/3.75/2.75pt against the natural 5.25 - up to 47.6% - to pull the
  next word on, ending flush at the text edge. Only ISOLATED single spaces
  set in an East Asian face compress (Times spaces and typed multi-space
  padding runs never do; cap modeled at 0.48). Trailing punctuation may hang
  past the edge (w:overflowPunct; the ":" of "zebeqo:" ends at 510.7 against
  a 505.35 edge).
- **Runs of >= 2 consecutive typed spaces lay at the EAST ASIAN space
  width** (5.25pt for SimSun) while isolated word spaces keep the Latin
  width - the 8-space padding run and ")  to  (" pairs in the p7 CSO-
  paragraph measure exactly so, ending the line at Word's x=472.5.
- **A uniformly lowered paragraph paints like unshifted text**: CSO- body
  runs all carry position -14 (-7pt) yet keep the 19.5pt pitch and the
  unlowered baselines - the common lowering is absorbed (mirror of the
  all-raised descent-reuse rule), both in plain lines and inside snapped
  lines.
- **Fit at the page bottom hangs grid leading**: an ordinary grid text
  line's fit extent is the raw font box (the last reference line's glyphs
  end 769.6 of 770 while its 15.6pt grid line overruns), and a grid-snapped
  object box does NOT reserve its paragraph spacing-after (Word keeps the
  (04) equation at bottom 766.6 of 770 with after=7.8 pending).
- Also: images/drawings on positioned runs now paint lowered/raised
  (`baseline - height - raise`), and the compat-15 justify pack extends to
  grid sections regardless of the document's compat mode.

Verification: all 8 pages <= 0.67% (mean 0.38), line counts 183 = Word's 183.
Gates: dense 12.63 (12.4 baseline, within run tolerance), chem 0.59 -> 0.31,
staging-eastasian 11.19 -> 7.95, parity-pictures/benchmark/sample/charstyles
0.00, compare-linebreaks canaries 6 x 0 mismatches.

## International text: RTL/bidi + East Asian (CJK) + docGrid

Added for the staging-bidi (1p), staging-eastasian (1p) and wild2-lit-yiddish-rtl
(215p) fixtures. Parse: rFonts w:eastAsia -> RunProps.fontEastAsia, w:cs ->
fontComplex, w:rtl -> RunProps.rtl, w:bidi (pPr) -> ParaProps.bidi, w:bidiVisual
(tblPr) -> TableProps.bidiVisual, w:docGrid -> SectionProps.docGridLinePitch.

- **CJK line breaking**: East Asian text has no spaces — every ideograph/kana is
  laid one em (= font size) wide and each inter-character boundary is a break
  opportunity (inline.ts pushCJK, breakAfter per char). Verified against Word's
  PDF: 42 CJK chars per line, x1=534pt = 72 + 42x11. Kinsoku: chars forbidden at
  line start (closing punctuation/small kana) or line end (opening brackets) are
  bound to their neighbour via breakAfter=false. jc=both does NOT stretch a
  CJK-only line (no spaces to distribute) — Word leaves it at natural width.
- **CJK font substitution & line pitch**: Word picks the CJK face by glyph
  coverage. A Japanese eastAsia font (MS Mincho) doesn't cover simplified
  Chinese, so Word falls back to Microsoft JhengHei (its PDF embeds it) with a
  much taller line box. Proxied with kana presence: a CJK run with no kana under
  a Japanese eastAsia font uses the Chinese profile. Measured 11pt line pitch
  (docDefaults line=259 -> x1.0792): MS Mincho 19.5pt/line (single 1.643em,
  ascent 1.364em); JhengHei 36pt/line (single 3.033em, ascent 2.27em). macOS
  substitutes: MS Mincho->Hiragino Mincho ProN, JhengHei->PingFang TC. Glyph
  WIDTH is 1em regardless of substitute; only the vertical metrics matter.
- **docGrid (type=lines/linesAndChars)**: sets a MINIMUM single-line font height
  = linePitch that the line-spacing multiplier is applied over (Latin lines in a
  CJK section grow to the grid; CJK fonts already exceed it). type="default"
  (grid defined but unused) does NOT snap — the yiddish book's docGrid is
  type=default, correctly ignored. FIRST-LINE reserve: Word drops the first line
  of a docGrid section ~4x linePitch below the top margin and suppresses that
  paragraph's spacing-before (measured: staging-eastasian's first heading
  baseline sits 4 pitches below the margin). This single reserve was the biggest
  eastasian win (85% -> 21% structural).
- **Bidi paragraph alignment**: w:bidi lines assemble logically then reorder to
  visual order (UAX#9 rule L2: reverse contiguous runs from the highest
  embedding level down; RTL runs get span.rtl so the renderer sets
  direction:rtl and the browser shapes Arabic/Hebrew within the box). Physical
  alignment FLIPS: OOXML jc="right" means "end", which in an RTL paragraph is the
  LEFT margin (measured: Word lays a bidi jc=right paragraph flush LEFT);
  jc="left"->right; absent->right (RTL start).
- **bidiVisual (RTL) table**: column order reverses (source col 0 lands at the
  right edge) AND the table hugs the right margin. Cell RTL text still uses the
  bidi paragraph swap (flush left within the cell) — right-aligning cell text
  measured WORSE.
- **Complex-script font**: a w:rtl run paints in the rFonts w:cs face (Arial for
  the bidi fixtures). Using it (vs the Latin fallback) dropped staging-bidi from
  12% -> 5% by matching Word's Arabic/Hebrew shaping/advances.

Results: staging-bidi 66.05% -> 5.24%, staging-eastasian 83.90% -> 20.90%
structural. Eastasian's residual is dominated by CJK glyph-shape decorrelation
(Hiragino/PingFang vs Word's embedded MS Mincho/JhengHei — unavailable on macOS;
raw pixel diff is only 4.6%) plus a few-pt vertical drift at Latin/heading
paragraph boundaries. wild2-lit-yiddish-rtl baselined at mean 76.65% over 215
pages: the RTL implementation renders it structurally correct (right-aligned,
justified, correct line breaks early), but per-glyph Times New Roman
decorrelation (mac vs Windows) plus sub-pt line-pitch drift diverge the
pagination over book length — the per-font line-advance calibration backlog,
compounded across hundreds of pages.
