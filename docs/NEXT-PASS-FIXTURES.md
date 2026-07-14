# Next-pass parity fixtures

Edge-case `.docx` fixtures built to pin behaviors the live benchmark bed does
not yet exercise. Each fixture is a minimal, structurally-valid OOXML package
(passes `scripts/validate-docx.py`), exported to a Word ground-truth PDF, and
baselined against the current DocxInWeb engine with
`DXW_PARITY_FAST=1 node scripts/parity-compare.mjs <name>`.

The docx sources and Word reference PDFs live under `fixtures-staging/` and
`parity/` (both gitignored); regenerate with the wave's generator script. The
severity numbers below are the current-engine baseline (Word vs web), so they
say where the engine is weakest, worst-first.

<!-- Wave 1 owns the section below this line; append new waves at the end. -->

## Wave 1

Generator: `scripts/generators/make-probe2-fixtures.py` (self-contained; no external
helper import). Fixtures are `probe2-<slug>.docx`; references are
`parity/probe2-<slug>-word.pdf`.

Build + validate + baseline:

```
python3 scripts/generators/make-probe2-fixtures.py                # all 12 fixtures
python3 scripts/validate-docx.py fixtures-staging/probe2-*.docx
# copy each into apps/demo/public/fixtures TEMPORARILY, then:
DXW_PARITY_FAST=1 node scripts/parity-compare.mjs probe2-<slug>
```

All 12 fixtures materialized in Word (confirmed page by page from the PDFs); none
needed hand-authoring. Baselined 2026-07-11 against `main`.

### Baseline results (worst-first)

Severity is the per-page structural mismatch % the parity harness reports (peak
page in bold).

| Fixture | Construct pinned | Peak severity (page) | Where the engine is losing |
|---|---|---|---|
| `probe2-picture-watermark` | VML header picture watermark (washout `gain`/`blacklevel` + a full-strength section) + rotated WordArt `t136` textpath; embedded PNG | **100% (p7, extra page)**, 94% (p6), 18–43% (p1–5) | Header `v:imagedata` painted at full opacity and top-anchored in the header flow, not washed-out / page-centred / behind text — which inflates the header into a 7th page the Word render doesn't have. Implement VML washout + `mso-position:center relative:margin` behind-text placement with zero flow contribution. |
| `probe2-ruby-vertical` | `w:ruby` (rt over base) at two sizes; `tbRl` + `btLr` wrapping cells; page-level `textDirection=tbRl` | **88% (p2)**, 35% (p1) | `w:textDirection` ignored — vertical-writing sections/cells are laid out horizontally (diff shows the tbRl page left-to-right); ruby rt-over-base stacking absent. Implement vertical writing mode + ruby layout. |
| `probe2-mixed-orientation` | portrait A4 → landscape Letter (1.5in margins) → portrait `w:vAlign=center`; per-section headers/footers; tables spanning breaks | **72% (p3)**, 0% (p1/p2) | Orientation + margin switches already at parity; only section `w:vAlign=center` is missing — content is pinned to the top instead of the vertical middle. Implement the section vertical-alignment pass. |
| `probe2-math-matrices` | OMML `m:m` 2×2/3×3 (fractions), `m:d`+`m:eqArr` piecewise, `m:acc`/`m:groupChr`, `m:rad` degree + nesting, `m:limLow`/`m:limUpp` | **49% (p1)** | These OMML node types aren't laid out: matrix grid (`m:mcs` column jc), stretched delimiters around arrays, accent/group-char positioning, radical degree/nesting, and under/over operator limits. |
| `probe2-run-borders` | `w:bdr` single/double/dotted/dashed/thick, ± `w:shd`; adjacent identical runs merge vs differing stay separate; a bordered run wrapping across lines | **31% (p1)** | Run border boxes are largely unpainted. Implement `w:bdr` boxes with the run's shading, coalescing consecutive identical-bdr runs into one box, and a box segment per line-fragment on wrap. |
| `probe2-arabic-rtl` | RTL Arabic w/ tashkeel, mixed Arabic/Latin/digit bidi runs, right-aligned tab line, `numFmt=arabicAbjad` RTL list | **13% (p1)** | Bidi ordering mostly works; residual is Arabic glyph shaping/joining metrics, tashkeel placement, and abjad list-marker glyphs. |
| `probe2-content-controls` | rich/plain/date/inline/nested/group SDTs, grey `showingPlcHdr`, table-row `w15:repeatingSection`, SDT-wrapped TOC w/ dot leaders | **9.4% (p1, weight)** | Content renders (drift class `weight` = ink-mass, not layout). Tune placeholder grey styling, repeating-row spacing, TOC leader fill. |
| `probe2-dropcaps-frames` | `framePr` drop caps (`drop` lines 3 & 2, `margin` lines 3) + two positioned bordered/shaded frames with wrap | **2.9% (p1)**, 0% (p2) | Near parity; minor drop-cap glyph metric/baseline tuning. Frame wrap already tracks. |
| `probe2-form-checkboxes` | legacy `FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN` (cached) + `w14:checkbox` SDTs, inside & outside table cells | **1.2% (p1)** | Near parity; checkbox glyph/box metric nits only. |
| `probe2-hyphenation` | `autoHyphenation` + zone + `consecutiveHyphenLimit` + `doNotHyphenateCaps`; justified/left; all-caps; soft hyphen + `noBreakHyphen`; narrow 2-col | **0.5% (p1)**, 0% (p2/p3) | Effectively solved — auto-hyphenation, caps suppression, and non-breaking hyphen all honoured. |
| `probe2-styleref-headers` | `STYLEREF "Heading 1"` (first-on-page) + `\l` (last-on-page) + `PAGE`; a chapter starting mid-page | 0.1% (p1), **0.3% (p2)** | Effectively solved — STYLEREF first/last-on-page resolution matches (p1 header: "Chapter One" vs `\l` "Chapter Three"). |
| `probe2-modern-template` | themed heading colours/fonts, SDT cover block, banded Grid Table (`tblLook` + `tblStylePr` + theme fills), accent pull-quote box; ships `theme1.xml` | **0.0%** | Solved — theme plumbing (colours, fonts, table banding, scheme fills) is at parity in combination. |

### Notes

- A **block** repeating-section SDT wrapping whole tables makes Word open the doc
  but reject `save as` (silent -1708) — the fixture uses the canonical **table-row**
  form (`w15:repeatingSectionItem` per `<w:tr>`). Separately, a stale open doc in
  the shared single-instance Word also causes spurious -1708 export failures; close
  the doc by name before retrying.
- `probe2-picture-watermark` is the only fixture whose engine page count diverges
  from Word (7 web vs 6 Word) — tracked as the peak finding above.

## Wave 2

Generator: `scripts/generators/make-probe3-fixtures.py` (imports helpers from
`scripts/generators/make-staging-fixtures.py`; run with **python3.12+** — the helper
library uses backslashes in f-strings). Fixtures are `probe3-<slug>.docx`;
references are `parity/probe3-<slug>-word.pdf`.

Build + validate + baseline:

```
python3.12 scripts/generators/make-probe3-fixtures.py           # all fixtures + robustness/
python3.12 scripts/validate-docx.py fixtures-staging/probe3-*.docx
# copy each into apps/demo/public/fixtures TEMPORARILY, then:
DXW_PARITY_FAST=1 node scripts/parity-compare.mjs probe3-<slug>
```

All 15 hand-authored fixtures + 1 LibreOffice-authored fixture materialized in
Word (confirmed by rendering the PDFs). Two required documented fallbacks:
tracked-changes markup and one WordArt warp (see notes).

### Baseline results (worst-first)

Severity is the per-page structural mismatch % the parity harness reports; the
worst page per fixture is bolded in the notes.

| Fixture | Construct pinned | Peak severity (page) | Where the engine is losing |
|---|---|---|---|
| `probe3-chargrid` | `docGrid type=charsAndLines` char pitch + line grid; `snapToGrid=0` contrast | **100% (p2, extra page)**, 34% (p1) | Char-cell padding + line-pitch snapping unimplemented → lines set too tall → an extra page the Word render doesn't have. |
| `probe3-columns-unequal` | 3 unequal `w:col` widths + separator rule + manual column breaks; a balanced 3-col section | **68% (p1)**, 1% (p2) | Multi-column section layout not applied — content flows full-width single-column instead of into columns. |
| `probe3-lo-provenance` | Genuinely LibreOffice-authored doc (HTML → `soffice --convert-to docx` → Word PDF) | **57% (p1)** | LibreOffice's own style/numbering/spacing conventions (list indents, table sizing, default spacing) drift from our style resolution. |
| `probe3-emoji` | Color emoji in body/heading/table; ZWJ family+couple, skin tones, flag pairs | **37% (p1)**, colour ΔE 18.3 | Color-emoji glyphs not rendered in color; ZWJ-cluster and regional-indicator advance widths differ (Word draws flag pairs as boxed letters — pin that). |
| `probe3-linked-textboxes` | Story flows box 1 → box 2 via `wps:linkedTxbx id/seq` | **33% (p1)** | Linked-textbox chaining unimplemented — box 2 stays empty and the story does not overflow between boxes. |
| `probe3-table-exotics` | Floating tables (`tblpPr`) overlapping (`tblOverlap`); old `tblCellSpacing`; diagonal `tl2br/tr2bl` borders; irregular merges | 21% (p1), **30% (p2)**, weight 60% | Floating/overlapping table positioning, separated-border cell spacing, and diagonal cell borders all missing. |
| `probe3-indic` | Devanagari + Tamil conjuncts/matras mixed with Latin | **17% (p1)** | Complex-script shaping (cluster formation, matra reordering) advance widths and break points differ. |
| `probe3-shape-autofit` | `normAutofit` (shrink), `spAutoFit` (grow), `noAutofit` (clip), same overfull text | **17% (p1)**, colour ΔE 15 | Text-box autofit not modeled: `spAutoFit` box growth and `normAutofit` re-scale not applied (box fills/sizes diverge). |
| `probe3-kashida` | Arabic `jc=distribute / lowKashida / mediumKashida / highKashida` | **14% (p1)** | Kashida letterform elongation + distribute justification for Arabic not implemented — justified Arabic sets differently. |
| `probe3-mirror-book` | `mirrorMargins` + gutter, distinct odd/even headers, page numbers at outer edges, 7 sections | 0.3% (p1), **9.6% (p2)** | Even-page margin mirroring / outer-edge page-number tab is slightly offset on verso pages. |
| `probe3-thai` | Thai (no spaces) + Lao, justified & left, narrow 2-col — dictionary line breaking | **3.8% (p1)**, 0% (p2) | Near parity: Thai break points mostly correct; only minor glyph-weight drift. |
| `probe3-tracked-changes` | `ins`/`del`/`moveFrom`/`moveTo`/`rPrChange`/`pPrChange`, deleted+inserted table rows | **2.8% (p1)** | Near parity against the **accepted/final** view (see markup note). Pins that our engine also resolves changes to final correctly. |
| `probe3-index-xrefs` | `XE`+`INDEX` cached 2-col result; `REF \h \r \p`; `SEQ \s` captions across chapters | **0.7% (p1)** | Near parity — cached field results render; deeper test is live recompute of INDEX/SEQ. |
| `probe3-text-effects` | Hidden text, outline/emboss/engrave/dstrike/shadow, highlight-over-shading, `fitText`, sub-font exact line, spacing+scale+kern | **0.6% (p1)** | Near parity — all effects render; hidden text correctly suppressed. |
| `probe3-field-switches` | `DATE/TIME \@`, `PAGE \* roman/ArabicDash`, `QUOTE \* Upper/Lower/FirstCap/Caps`, MERGEFORMAT vs CHARFORMAT, NUMPAGES/SECTIONPAGES, roman→arabic restart | **0.3% (p1)**; p2 line 75% | Structurally near-parity, but Word **recomputes DATE/TIME/PAGE on open** (renders today's date, not the cached value) — our engine shows the cached result, so text diverges. Decide cache-vs-recompute policy. |
| `probe3-wordart-warps` | `prstTxWarp`: textArchUp, textWave1, textChevron, textCirclePour | **0.2% (p1)**, weight 22% | Warp geometry not applied (warped glyphs are small ink area, so structural score stays low) — the shapes render but unwarped. |

### Notes / documented fallbacks

- **Tracked-changes markup does not survive Word's PDF export.** Word for Mac's
  `save as … PDF` always writes the *final/accepted* view. Setting the
  AppleScript levers `document.print revisions`, `view.show revisions and
  comments`, and `view.revisions view` (all confirmed against Word's `sdef`)
  does **not** inject strikethrough/underline markup into the exported PDF — the
  only markup path is the OS Print panel, which is not cleanly scriptable. Per
  plan, the committed reference is the clean/accepted view, and
  `scripts/word-parity-markup.sh` is provided (it sets those levers) so the
  markup export can be re-attempted if a future Word build honors them. The
  fixture therefore pins **accepted-view** rendering (deletions removed,
  insertions kept, moved text at destination, deleted row gone, inserted row
  present, format change applied).
- **`textArchUp` fell back to flat text** in Word's export; `textWave1`,
  `textChevron`, and `textCirclePour` all warped strongly. The arch case is a
  data point, not a broken fixture.
- **`normAutofit` was not honored on import** — Word clipped the overfull box at
  its fixed height (same as `noAutofit`) rather than applying the stored
  `fontScale`, while `spAutoFit` clearly grew its box. So the fixture actually
  pins two visible behaviors (grow vs clip), and that Word ignores a stored
  `normAutofit` scale without a layout recalc.
- **Regional-indicator flag pairs render as boxed letters** ("US", "JP") in
  Word on this machine, not flag images — that is Word's genuine behavior and a
  real parity target.

### Robustness crash-test set (NOT parity)

`fixtures-staging/robustness/` holds four deliberately damaged `.docx` for the
e2e loader (graceful-degradation, not pixel parity — Word refuses/repairs them,
so there is no ground truth):

- `damaged-missing-styles.docx` — `document.xml.rels` points at `styles.xml` but the part is absent.
- `damaged-unknown-elements.docx` — unknown `w:`-namespace and foreign-namespace elements sprinkled through the body.
- `damaged-2mb-paragraph.docx` — one paragraph with a single ~2 MB text run (perf/memory).
- `damaged-truncated.docx` — a valid package with its last 400 bytes chopped (broken central directory).
