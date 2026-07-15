# How WordInWeb reaches Word parity

WordInWeb renders `.docx` files in the browser and edits them. No server, no
conversion to HTML flow. It runs a layout engine that does what Word does: parse
the OOXML, break lines and paginate with real font metrics, and place every
glyph itself.

This is a note on how it got close to desktop Word, and the loop that made
"close" a number you can drive to zero.

## The layout engine

A `.docx` file stores content and formatting intent. It stores no positions.
Word decides where every line and page break lands, so any renderer has to make
the same decisions or accept being approximate.

Most viewers translate the document to flowing HTML and let the browser lay it
out. That is fast to build, but the browser's layout is not Word's, and you
cannot ask it where a page ends. Pagination, page numbers, and header geometry
are out of reach.

WordInWeb reimplements the layout engine and uses the browser only to paint.

```
bytes → Package → DocxDocument (typed model + retained XML) → layoutDocument()
      → LaidOutPage[] of {text|rect|edge|image} → renderToDom() → pixels
                                                → DocxEditor (mutates XML) ↩
```

Parse OOXML into a typed model. Run a canvas-measured line breaker and paginator
that resolve `PAGE`/`NUMPAGES` per page. Emit absolutely-positioned primitives:
`{x, baseline, font}` text items, rects, edges, images. The DOM renderer maps
each primitive to one absolutely-positioned element, so the browser reflows
nothing. Editing keeps the XML tree as the source of truth: commands mutate the
retained XML, the model is re-derived, layout re-runs, and `save()` re-serializes
only the parts it models so everything else survives byte-for-byte.

That architecture is what makes parity reachable. Every position comes from the
engine and the browser adds none, so there is a single number to drive toward:
the pixel difference between our render and Word's.

## The training loop

The reference is desktop Word. For every fixture:

1. Word exports the `.docx` to PDF.
2. WordInWeb renders the same file headless.
3. Both rasterize to PNG at the same DPI and compare page by page.

The diff produces a score and a diff image. That gives an agent a closed loop
with a definite end: render, diff, read the diff image, form a hypothesis about
which Word rule is wrong, change the engine, re-measure. The number either drops
or it doesn't. No human has to judge each attempt, so the loop can iterate on the
specifics until the score reaches zero.

## The metric

Counting differing pixels does not work. Chrome and Word antialias glyph strokes
differently, so two identical layouts still differ along every stroke edge. The
headline metric is structural severity: after registering one global page offset,
it is the fraction of binary ink with no counterpart within a small local
tolerance, plus ink that only matches after a one-line vertical shift beyond the
calibrated noise floor. Antialiasing noise and sub-pixel line shifts do not
count. Each page also gets a `driftClass` (clean / alignment / weight / colour /
structural) so a weight-only or colour-only difference is named instead of folded
into the score.

Two rules keep the harness honest. Never calibrate against your own measurements
or pdfminer word extents; build a probe document, export it through Word, and
read the geometry back out of the PDF. And a persistent raster cache exports each
Word PDF once, so a full run over 1,154 pages takes seconds, not minutes.

Fixtures are grouped into capability areas (general word processing, complex
tables, math, other languages, formatting, graphics, real-world documents), each
with its own subscore, so a weak area is visible instead of buried. A separate
LibreOffice tab tracks LibreOffice-authored references on their own axis so
cross-suite differences never move the Word number.

```bash
node scripts/parity-parallel.mjs                     # full run → parity/out/results.json
DXW_PARITY_FAST=1 node scripts/parity-parallel.mjs   # skip slow appearance passes
node scripts/parity-render-report.mjs                # rebuild parity/out/report.html
```

## One fix the loop found: Devanagari spaces

`probe3-indic` p1 sat at 17.26%. Devanagari lines wrapped about 57px early and
the shift cascaded down the page. At the glyph level (Chrome canvas plus
pdfplumber against Word's PDF) Chrome shaped the Mangal conjuncts correctly and
per-word advances matched Word to under 0.1px. The error was entirely the spaces.

Mangal's U+0020 advance is 0.5em; Calibri's is 0.226em. At 11pt that is 5.5px
versus 2.49px. Word resolves an ASCII space to the `w:ascii` face, not the run's
complex-script `w:cs` face, but the engine had routed the whole run to Mangal.
Every inter-word gap was about 2.75px too wide, and 20 spaces per line is 55px of
overshoot, enough to wrap every line early.

The rule is general: OOXML font resolution is per character class (ASCII→`w:ascii`,
CJK→`w:eastAsia`, complex→`w:cs`), and spaces are ASCII. Routing spaces back to
the ascii face took the page from 17.26% to 5.76%. Scaling the Tamil Latha
fallback to match Word's Vijaya metrics took it to 1.30%, where it now sits
against a licensed font we cannot ship.

## The same loop for editing speed

Rendering parity is measured against pixels; editing speed is measured against a
keystroke-timing harness with per-keystroke budgets. Same shape: a testable
target the change either hits or misses.

The naive pipeline re-ran everything on every keystroke: full model reparse, full
document layout, full DOM teardown and rebuild. On a 419-page contract,
mid-document, that was 5,005 ms per keystroke (layout 1,587 ms, render
2,276 ms). Profiling ruled out the easy suspects first: a warm text measurer
saved about 15%, and math layout was 1.3 ms across 38 equations. The cost was the
line-break and pagination walk over every page, plus the DOM rebuild.

The fix is incremental, and pinned:

1. **Incremental pagination.** `layoutDocument(doc, {prev})` reuses page objects
   for the unchanged prefix and suffix and relays only the changed middle.
   Mid-doc layout went 1,158 ms → 151 ms, byte-identical to a full relayout.
2. **Incremental DOM reuse.** The renderer adopts the DOM of pages whose layout
   is unchanged and splices only the changed-page window into the existing root.
   Rebuilding a fresh root re-appended 419 page elements and forced about 800 ms
   of hidden reflow. Render went 1,019 ms → 53 ms.
3. **Dirty-block hints.** The editor tells layout which block the caret is in, so
   the incremental pass hashes that block and two neighbors instead of scanning
   all of them.

`incr-equiv.test.ts` asserts incremental layout is byte-identical to full layout
across fixtures and edit positions, because parity is a read-only measurement and
cannot catch edit-path bugs. Medians (headless, 20 keystrokes): the 419-page
contract went 5,005 ms → ~250 ms; a 17-page dense-math page 212 ms → ~72 ms.

The recurring lesson: every "still laggy" report was a class of document that
quietly disabled reuse. Comments disabled page adoption until it was fixed;
images and the full-root reflow did the same. You only find those against real
documents.

## Where it stands

Certified run, sha `92077b9`, metric `ink-dilate-line-v5`:

| Metric | Value |
| --- | --- |
| Word pages measured | 1,154 across 91 fixtures |
| Mean structural severity | 0.026% |
| Worst single page | 3.95% (`probe3-thai` p1) |
| Pages at or above 1% | 2 (thai 3.95%, indic 1.30%) |

Both pages above 1% are a licensed font we cannot legally ship: Thai DokChampa,
and Tamil Vijaya approximated by a scaled Latha. The rest is within the
antialiasing floor between two rasterizers drawing the same layout.

Reproduce with `node scripts/parity-parallel.mjs`, which writes
`parity/out/results.json`.
