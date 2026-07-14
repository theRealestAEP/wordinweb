# How we got DocxInWeb close to Word parity

DocxInWeb renders `.docx` files natively in the browser — no server round-trip,
no image conversion, no "download to view in Word." The bar we set was not
"looks about right." It was **pixel parity with desktop Microsoft Word**, page
for page, measured against Word's own output.

This is the story of how we got there: the architectural bet that made it
possible, the measurement rig that kept us honest, and the long tail of
"irreducible" pages that turned out to be bugs. As of the certified run
(sha `92077b9`, metric `ink-dilate-line-v5`):

| Metric | Value |
| --- | --- |
| Word-authored pages measured | **1,154** across 91 fixtures |
| Mean structural severity | **0.026%** |
| Worst single page | **3.95%** (`probe3-thai` p1) |
| Pages at or above 1% severity | **2** (thai 3.95%, indic 1.30%) |
| Pages with no named drift (`driftClass = clean`) | **1,053 / 1,154** |

Methodology lives in [`docs/EVALS.md`](EVALS.md); the non-obvious Word behaviors
we reverse-engineered are catalogued in [`docs/DISCOVERIES.md`](DISCOVERIES.md).
This post is the narrative that connects them.

---

## 1. The bet: a real layout engine, not a docx-to-HTML mapping

A `.docx` file contains no positions. It carries content and formatting
*intent* — styles, spacing rules, section geometry — and leaves every line
break and page break to whoever lays it out. That single fact forces a choice,
and the choice determines your ceiling.

- **Translate to flowing HTML/CSS** (the docx-preview / mammoth approach). Fast
  to build. But the browser's layout algorithm is not Word's, and you cannot
  ask the browser where a page ends. Pagination, page numbers, header geometry,
  divider-line placement — all unachievable. You are permanently approximate.
- **Reimplement the layout engine** and use the browser only as a display
  surface.

DocxInWeb is the second. The whole system hangs off one contract:

> Layout produces absolutely-positioned primitives. Nothing downstream makes a
> spatial decision.

```
bytes → Package → DocxDocument (model + retained XML) → layoutDocument()
      → LaidOutPage[] of {text|rect|edge|image} → renderToDom() → pixels
```

We parse OOXML into a typed model, run our own canvas-measured line breaker and
paginator, resolve `PAGE`/`NUMPAGES` during pagination, and emit
`{x, baseline, font}` text items and positioned rects/edges. The DOM renderer
maps each primitive to one absolutely-positioned element and the browser
reflows *nothing*. What the engine computed is exactly what you see.

The cost is real: we own line breaking, justification, widow/orphan control,
table autofit, header/body-top geometry, footnote reserve, column balancing,
bidi, and font metrics. Every one of those is a place Word can surprise you. The
payoff is that pixel parity is even *possible* — with an HTML-flow mapping it is
not on the table at all.

## 2. Ground truth: Word is the reference, and raw pixels are useless

The reference is not a spec or another renderer. It is desktop Microsoft Word
itself. For every fixture:

1. Word exports the `.docx` to PDF — the ground truth.
2. DocxInWeb renders the same file headless.
3. Both are rasterized to PNG at the same DPI and compared **page by page**.

The obvious metric — count differing pixels — is worthless. Chrome and Word
antialias glyphs differently, so two *identical* layouts differ on a large
fraction of pixels along every stroke edge. Our headline metric is **structural
severity**: after registering one global page offset, it is the fraction of
binary ink with no counterpart within a small local tolerance, plus ink that
only matches after a one-line vertical shift beyond the calibrated noise floor.
Pure antialiasing noise and sub-pixel line shifts do not count — only genuine
missing/extra/misplaced ink does. Each page also gets a `driftClass` (clean /
alignment / weight / colour / structural) so a weight-only or colour-only
difference is *named*, not silently folded into the score.

The harness had to be fast enough to run constantly and trustworthy enough to
believe:

- **Parallel sharding** (`scripts/parity-parallel.mjs`) fans 1,154 pages across
  workers.
- A **persistent reference raster cache** — Word PDFs are exported *once* per
  fixture and rasterized once; Word is never invoked per comparison run. Runs
  that used to take many minutes became seconds.
- **Provenance tabs**: the **Word** tab is the parity target. A separate
  **LibreOffice** tab tracks LibreOffice-authored references on their own drift
  axis so cross-suite differences never pollute the Word headline.

One rule underwrites all of it, learned the hard way: **never calibrate against
your own measurements or against pdfminer word extents.** Build a probe
document, export it *through Word*, and read the geometry back out of the PDF.
Word-on-Mac's PDF writer even spreads justification residue into intra-word
glyph adjustments (±5 milli-em of pure noise) — an entire early justify
calibration was fitted to that noise before we caught it.

## 3. "It's a rasterization floor" is almost always false

The most important cultural lesson of the project. When a page sits at 5%, 12%,
17% severity, the tempting story is "that's just Chrome vs Word rasterizing
complex glyphs differently — irreducible." Nearly every time we told ourselves
that above ~1%, it was a real bug hiding behind the excuse. The complex-script
suite is the cleanest example.

**Indic: 17.26% → 5.76% → 1.30%.** `probe3-indic` p1 was stuck at 17.26%.
Devanagari lines wrapped ~57px early and cascaded a horizontal shift down the
page. We proved at the glyph level (Chrome canvas + pdfplumber vs Word's PDF)
that Chrome shapes Mangal conjuncts perfectly and per-word advances matched Word
to <0.1px. The entire error was the **spaces**: Mangal's U+0020 advance is
0.5em versus Calibri's 0.226em. Word resolves an ASCII space to the `w:ascii`
face, not the complex-script `w:cs` face — we had routed the whole run to
Mangal, doubling every inter-word gap (20 spaces × 2.75px ≈ 55px/line). The rule
is general: OOXML font resolution is per-character-class (ASCII→`w:ascii`,
CJK→`w:eastAsia`, complex→`w:cs`), and spaces are ASCII. That took it to 5.76%;
scaling the Tamil Latha fallback to match Word's Vijaya metrics took the last
step to 1.30%.

**Kashida: 14.02% → 0.95%.** Arabic kashida justification elongates the joins
between letters rather than stretching spaces. Modeling the per-flavor
(`low`/`medium`/`high`) join elongation made our line counts match Word's — the
severity was never rasterization, it was that we were breaking lines in the
wrong places.

**Arabic RTL: 12.98% → 0.00%.** A UAX#9 bidi embedding pass, right-edge wrap for
explicit left-tabs in RTL, and the `arabicAbjad`/`arabicAlpha` list-marker gap.
Zero. A page that had been "obviously a shaping floor" became pixel-clean.

And the one honest floor that remains: **Thai/Lao, 3.95%.** The reference uses
DokChampa, a licensed Microsoft font we cannot ship. We bundle an OFL substitute
(Noto Sans Lao Looped) for the Lao script, but the metric still penalizes the
render against a typeface we are not allowed to distribute. That is the worst
page in the entire 1,154 — and it is a licensing constraint, not a layout bug.
The discipline the suite enforces: **no floor claim without glyph-level
receipts.** Every "irreducible" label has a probe behind it or it doesn't count.

## 4. Word's rules are discoverable, but weird

Word's behavior is not documented where you'd hope, but it *is* deterministic,
and a probe document exported through Word will tell you exactly what it does.
A few of the ones that cost the most time:

**Spacing collapses across section breaks, but not always.** Adjacent paragraph
spacing doesn't add — Word takes the max of one paragraph's space-after and the
next's space-before. Fine. But a page break inside a section drops the following
paragraph's space-before, *while a new section's opening paragraph does not* —
it follows a cross-section "carry-remainder" rule
(`max(before, carriedAfter) − carriedAfter`). Conflating the two put a heading
15pt too high and read as 38% on one page of a multi-column template.

**There are two different soft hyphens, and conflating them is catastrophic.** A
raw `U+00AD` character sitting in a `w:t` run paints as a *visible* hyphen in
Word. A `<w:softHyphen/>` *element* is a conditional hyphen — invisible unless
the line actually breaks there. They look identical in a naive reading and mean
opposite things. Treating the element like the character exploded an IEEE
two-column template to 86% severity before we bisected it. The conversion now
happens precisely where `w:t` text is parsed.

**Fields recompute on open, so parity has to freeze the clock.** A `DATE` field
resolves to *today* every time the document opens, which means our render and
Word's PDF disagree by construction unless we pin our page clock to the
reference PDF's `CreationDate`. `REF`/`PAGEREF` fields likewise carry stale
caches in the file that Word recomputes on open — you cannot trust the cached
value in the XML; you have to recompute what Word would.

**Margin drop caps behave differently by anchor.** A `dropCap="margin"` cap
only lifts into the margin when its frame is `hAnchor="page"`; a text-anchored
margin cap actually behaves like an in-text drop cap. Gating on the anchor was
the difference between a clean page and a regression.

**NBSP glues words that a normal space would break.** A space run whose next
word begins with a non-breaking space is not a break opportunity — Word moves
the whole glued unit (`"of $ [nbsp…] (blank)"`) to the next line together. Miss
it and you fit one extra line per fill-in-blank table, cascading into
`keepNext`/`cantSplit` flips at every page bottom downstream.

Each of these was pinned with `scripts/make-*-probe*.py` — a synthetic document
built to isolate one behavior, exported through Word, and read back from the
PDF. That loop is the single most valuable tool in the repo.

## 5. Making editing fast: 5,005 ms → ~250 ms per keystroke

Rendering parity was the goal; editing had to not be miserable. The naive
pipeline re-ran everything on every keystroke: `doc.refresh()` (full model
reparse) + `layoutDocument()` (full document layout) + `renderToDom()` (full DOM
teardown and rebuild). On a 419-page contract, mid-document, that was **5,005 ms
per keystroke** — layout 1,587 ms, render 2,276 ms.

Profiling killed the easy theories first. Layout was not dominated by text
measurement (a persistent warm measurer cut it only ~15%) or by math
(`layoutMath` was 1.3 ms across 38 equations). It was the per-paragraph
line-break + pagination walk over *all* pages. The fix is architectural, not a
cache:

1. **Incremental DOM reuse.** `renderToDom` adopts the DOM of pages whose layout
   is unchanged via a common prefix + suffix page-equality diff. A keystroke
   rebuilds ~1 page; 418 of 419 are reused.
2. **Incremental pagination.** `layoutDocument(doc, {prev})` reuses page objects
   for the unchanged prefix and (on re-convergence) suffix, relaying only the
   changed middle. NIH mid-doc layout went 1,158 ms → 151 ms, **byte-identical**
   to a full relayout.
3. **In-place root splice.** When mounted, reuse the previous DOM root and
   `replaceChild()` only the changed-page window. Building a fresh root and
   re-appending 419 page elements forced the browser to reflow every page
   (~800 ms of hidden reflow). Render now scales with pages *changed*, not pages
   *total*: NIH render 1,019 ms → 53 ms.
4. **Dirty-block hints.** The editor tells layout which block the caret is in, so
   the incremental pass hashes that block plus two neighbors instead of scanning
   all of them.

The correctness net matters as much as the speed. `incr-equiv.test.ts` asserts
that incremental layout is **byte-identical** to full layout (ignoring
editor-only `src` refs) for every fixture across multiple edit positions —
because parity is a read-only measurement and cannot catch edit-path bugs.

The recurring lesson: every "still laggy" report from a real user was a
*reuse-disabled class* we hadn't found yet. **Comments** silently disabled page
adoption (a 9-page commented letter rebuilt every page per keystroke: 81 ms →
28 ms once fixed). **Images** and the full-root reflow each did the same. Image
*decoding* — the thing everyone assumes is the cost — never was; it's cached per
document. The bottleneck is always whatever quietly forces a full rebuild, and
you only find it against real documents.

Final medians (headless, 20 keystrokes): the 419-page contract went **5,005 ms
→ ~250 ms**; a 17-page dense-math page **212 ms → ~72 ms**.

## 6. Process notes worth stealing

- **Fence tests around every fix.** The fixture most related to a change belongs
  in that change's regression gate. We learned this from a drop-cap fix that
  passed its own probe and regressed `parity2-dropcap` — which wasn't in the
  fence.
- **Candidate vs. accepted artifacts.** A run is a candidate until it's compared
  against the accepted baseline and promoted. The certified `results.json`
  carries `outcome: "accepted"` explicitly.
- **The dashboard can't see regressions until you re-measure.** A fixture can
  break and the report won't show it until that page is re-run. After any
  layout-wide merge we spot-run the whole real-world `wild2-*` corpus, not just
  the touched fixtures' fences — because a coordinated metric change can regress
  a page nobody edited.
- **Coordinated recalibration beats piecemeal fixes.** Chasing individual math
  constructs to zero *regressed* the matrix fixture under single-offset
  registration; recalibrating the whole family together took it to 0.00. Local
  optima fight each other.
- **No floor claim without receipts.** Repeated in section 3 because it's the
  one that mattered most. "It's just rasterization" is a hypothesis, not a
  conclusion, until a glyph-level probe says so.
- **Worktree hazard, documented so it stops biting.** A missing `fonts-local`
  symlink makes Vite serve fallback fonts and inflates *every* score — one
  branch saw `parity-lists` at 14% when the truth was 0%. Symlink the fonts and
  fixtures before trusting any number from a fresh worktree.

---

## Where it stands

1,154 Word-authored pages, mean structural severity **0.026%**, exactly two
pages above 1% — and both of those are a font we can't legally ship and a Tamil
fallback scaled to a font we can't legally ship. The rest is within the
antialiasing noise floor between two different rasterizers drawing the same
layout.

The engine still has open frontiers — figure anchoring in complex two-column
papers, a handful of table conditional-formatting details, per-font
sub-quarter-point vertical calibration — and each is tracked with the same
discipline: a probe, a receipt, a fence. If you want to reproduce the numbers,
`node scripts/parity-parallel.mjs` writes `parity/out/results.json`; see
[`docs/EVALS.md`](EVALS.md) for the full methodology.
