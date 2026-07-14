# Known limitations

An honest inventory for the first release. Every claim below is derived from the
code, the test suite, or the certified parity run — not aspiration. Sources are
cited inline (`file:line`, a fixture grep, or a `results.json` number).

The **viewer** is the mature surface: 1,154 Word pages compared page-by-page
against desktop Microsoft Word, mean structural severity **0.026%**, worst
**3.95%** (`parity/out/results.json`, git `92077b9`, `isFullRun: true`). The
**editor** is newer and opt-in (`editable` flag); its supported operations are
exactly the methods on `DocxViewApi` (`packages/react/src/index.tsx:58`). This
document catalogs where each falls short.

Difficulty tags: **quick-win** (wiring/UI only, engine already capable),
**planned** (real work, no architectural blocker), **architectural** (needs a
design change or is bounded by an external constraint).

---

## Editing: not yet supported

Eight gaps. Each is a capability a user could reasonably expect that the editor
does not provide today.

1. **Endnote insertion** — *planned.* Only footnotes can be authored
   (`insertFootnote`, `packages/core/src/edit/notes.ts:18`); there is no
   `insertEndnote`. Endnotes **render** correctly (model + layout support
   `w:endnoteReference`, `packages/core/src/model.ts:602-609`) and 24 fixtures
   exercise them, but the toolbar/API cannot create one.

2. **Structured equation editing** — *architectural.* Clicking an equation opens
   a linear-math popover (`openMathEditor`, `packages/core/src/edit/editor.ts:2064`).
   Equations that round-trip losslessly through the linear form are editable;
   matrices, n-ary / limit over-under groups, accents, and equation arrays open
   **read-only** so a stray edit cannot silently rewrite their OMML
   (`isLinearSafe`, `editor.ts:2060-2098`). You cannot type directly into an
   inline equation; you edit the linear string and re-emit OMML.

3. **Rich / image paste** — *planned.* Paste is **plain-text only**: the handler
   reads `clipboardData.getData("text/plain")` and ignores `text/html`,
   RTF, and image blobs (`packages/core/src/edit/editor.ts:646-651`). Copy and
   cut likewise emit only `text/plain` (`editor.ts:629-640`), so pasting between
   two points in the document drops all run formatting and styles. (Dragging an
   image **file** in from the OS does work — `onDrop`, `editor.ts:946`.)

4. **Suggesting mode tracks text only** — *architectural.* In suggesting mode,
   typed text and deletions record as `w:ins`/`w:del` via the suggestion-aware
   cores (`insertSuggestedText` / `deleteSuggestedRange`,
   `packages/core/src/edit/suggest.ts`). But formatting, tables, lists,
   paragraph styles, images, and page layout have **no suggesting branch**
   (`grep -n suggest packages/core/src/edit/{commands,tables,blocks,lists}.ts`
   is empty) — they mutate the document directly even while suggesting is on.
   Word records these as `rPrChange`/`pPrChange`/`tblPrChange`; we do not.

5. **Image cropping** — *planned.* Cropped images **render** (the `<img>` sits in
   a crop viewport, `editor.ts:1653`), but the editor exposes only resize
   (corner drag), reposition, wrap mode, alt-text, and blip replacement
   (`packages/core/src/edit/images.ts`: `setImageWrap`, `setFloatingPosition`,
   `setImageAltText`, `replaceImageBlip`). There are no crop handles.

6. **WordArt / watermark creation and styling** — *planned.* Only an **existing**
   VML text watermark can be changed, and only its string, rotation, and opacity
   (`setWordArtText` / `setWordArtRotation` / `setWordArtOpacity` / `deleteWatermark`,
   `packages/core/src/edit/watermark.ts`). There is no "insert watermark", no
   color or font control, and no editing of modern DrawingML WordArt.

7. **Field insertion beyond page numbers** — *planned.* The only field the editor
   inserts is a page-number field (`insertPageField`,
   `packages/core/src/edit/fields.ts:24`). Tables of contents, `DATE`, `REF` /
   cross-references, and other fields render if present in the file but cannot be
   inserted or updated from the editor.

8. **Charts** — *architectural.* No chart editing, and chart rendering is itself
   unsupported (see Rendering → Untested below). A `chartSpace` in the document
   has no code path.

Table structural ops that **are** supported, for contrast: insert/delete
row+column, merge right/down, split cell, delete table, cell shading, cell
vertical align (`TableOp`, `packages/core/src/edit/tables.ts:13`), including on
nested tables (`e2e/table-exotic-edit.spec.ts`). Not supported: per-edge cell
border editing, table-style application, numeric column-width entry (grip-drag
only).

---

## Editing: supported but shallow (quick wins)

These work, but the demo toolbar (`packages/react/src/toolbar.tsx`) hardwires a
narrow set of choices where the underlying command accepts more. Each is a
UI-only change.

1. **Page-border color is blue-only** — the user's headline example. The Layout
   dropdown offers four presets — none / thin box / thick box / **"Blue box"** —
   and blue (`4472C4`) is the *only* colored option
   (`toolbar.tsx:663-668`). The command itself accepts any hex color
   (`blocks.ts:198-210`), so exposing a color picker is trivial. The command
   also always writes `val="single"` on all four edges (`blocks.ts:204`), so
   line style (double/dashed/art borders) and per-edge control are the deeper
   follow-ups.

2. **Highlight: 5 colors** — `HIGHLIGHTS` offers yellow/green/cyan/magenta/
   light-gray (`toolbar.tsx:40-46`); Word has ~15 named highlight colors.
   (Text color, by contrast, uses a full native color picker — `toolbar.tsx:944`.)

3. **Font list is a fixed 35-family probe list** — `FONT_CANDIDATES`
   (`toolbar.tsx:7-13`) filtered by a canvas width probe. There is no system-font
   enumeration, so a font installed on the machine but absent from the list is
   unreachable unless it happens to be the current selection.

4. **Font sizes are a fixed preset list** — `SIZES = [8…48]` as a `<select>`
   (`toolbar.tsx:38`); no arbitrary point-size entry.

5. **Margins: 3 presets** — Normal / Narrow / Wide only (`toolbar.tsx:651`); no
   custom margins. `setPageLayout` accepts arbitrary inch values.

6. **Page size: Letter / Legal / A4** (`toolbar.tsx:657`); no custom dimensions.

7. **Columns: 1/2/3, equal width** (`toolbar.tsx:662`); the command hardcodes a
   0.5in gutter (`blocks.ts:216`) and offers no custom widths or spacing.

8. **Line spacing + paragraph spacing presets** — spacing menu offers
   1 / 1.15 / 1.5 / 2 and a fixed 10pt add/remove for space before/after
   (`toolbar.tsx:973-989`); no custom point values.

9. **Table cell fill: 5 swatches** (`toolbar.tsx:512`) + none; no custom color.

10. **Line numbers: count-by 5 or 10** presets only (`toolbar.tsx:669-683`).

---

## Rendering: known gaps

Across the certified full run, **22 Word pages score ≥ 0.5% structural
severity** (out of 1,154; every other page is under 0.5%, most under 0.05%).
They fall into two buckets.

### Accepted floors (not bugs — external constraints)

These are irreducible without shipping licensed fonts or matching Word's
rasterizer. Evidence lives in the fixtures and in `parity-metric-floors` memory.

| Page | Sev % | Why it cannot be closed |
| --- | ---: | --- |
| `probe3-thai` p1 | **3.95** | Word uses licensed **DokChampa**, which we cannot bundle. Lao renders the bundled OFL **Noto Sans Lao Looped** — style-correct, not glyph-identical; the metric penalizes the unshippable reference. |
| `probe3-indic` p1 | **1.30** | Tamil **Vijaya** (Word) vs bundled **Latha**, scaled. Different licensed typeface. |
| `probe3-kashida` p1 | **0.95** | Arabic justification: line counts match Word; residual is concentrated kashida (elongation) join placement — rasterization-class. |
| `probe3-emoji` p1 | **0.57** | **Apple Color Emoji vs Segoe UI Emoji** artwork. Explicitly user-accepted ("those will just be different, that's fine"). Never pixel-matches. |

Also accepted (below 0.5% now but same class): the yiddish RTL body pages are
pure stroke-rasterization difference of the *same embedded font* on both sides
(line breaks and glyph positions are pixel-exact) — irreducible without Word's
rasterizer.

### TODO (real residuals worth chasing)

| Page(s) | Sev % | Explanation / TODO |
| --- | ---: | --- |
| `parity2-equations` p1 | **0.82** | OMML raster-offset (equation baseline lead). Tracked as an open math residual. |
| `probe2-picture-watermark` p6 | **0.93** | Picture-watermark placement/opacity residual on one page. |
| `wild2-math-omml-dense` p8, p13 | 0.55, 0.54 | Compound inline n-ary math (∑∫ wrapping fractions) shows bidirectional ±2px row-pitch drift; a blanket ascent nudge fixes one sign and worsens the other — needs its own Word probe (see `parity-metric-floors`). |
| `wild2-legal-nih-contract` p35/103/316/317/414 | 0.51–0.69 | Recurring sub-1% residuals on a 419-page real-world contract; not yet individually diagnosed. |
| `parity-rowsplit` p1, p2 | 0.68, 0.67 | Table row-split-across-page-boundary residual. |
| `wild2-sci-ieee-2col` p4 | 0.79 | Two-column science layout; 11.7% raw pixel diff localizes to sub-line drift. |
| `wild2-math-eq-as-images` p2, p5 | 0.56, 0.57 | Equations shipped as raster images — sub-pixel resampling texture. |
| `wild2-sci-chem-omml` p4 | 0.55 | Chemistry OMML page residual. |
| `wild-doerfp` p8 | 0.82 | Real-world document, one-page residual, undiagnosed. |
| `staging-tblextreme` p1 | 0.52 | Extreme-table stress page. |
| `wild2-med-phase23-protocol` p18 | 0.50 | Page carrying the corpus's only **SmartArt** diagram (see below). |

### Rendering weaknesses / missing assets (not per-page scores)

- **Missing Office fonts.** The engine measures on canvas metrics, so *layout*
  matches Word even when a font is absent (metric-compatible substitutes:
  Calibri→Carlito, Cambria→Caladea, DokChampa→Noto Sans Lao Looped —
  `apps/demo/src/main.tsx:2-13`). But *paint* uses whatever the browser has. The
  viewer surfaces this: `onMissingFonts` fires and the demo shows a warning
  banner (`apps/demo/src/main.tsx:388`). Real Office faces (Cambria Math, real
  Calibri/Times/Arial, CJK families) load dev-only from `/fonts-local/` and are
  git-ignored — they are **not** shipped.
- **Rasterizer antialiasing floor.** Chrome and Word antialias glyph strokes
  differently; on pages where geometry is pixel-exact this is the entire
  remaining residual (README "Rendering parity"). Not closable in a browser.
- **Raster-image resampling.** Documents that embed equations/figures as bitmaps
  (`wild2-math-eq-as-images`, `wild2-sci-chem-omml`) carry a sub-pixel resample
  texture difference that is not a layout target.

---

## Untested

Areas with **zero or near-zero coverage**. "Untested" ≠ "broken" — it means we
have no fixture or spec proving parity, so behavior is unverified.

### Editing round-trip is unimplemented

The most important gap. The plan to prove *edit → `save()` → re-open in Word
matches* is designed in `docs/NON-FIXTURE-AXES-PLAN.md` (Axis 1) but **not
built**: there is no `e2e/edit-roundtrip.spec.ts`, no `e2e/edit-scripts/`, and no
`scripts/edit-roundtrip.mjs` (all confirmed absent). Every status row in that
plan reads *Incomplete*. So while individual edit operations have in-browser
specs (`e2e/*.spec.ts`), **no test verifies that a saved `.docx` re-opens
faithfully in Word.** Serialization bugs would not be caught today.

The **performance budget** axis (Axis 2) is likewise only partially present:
`e2e/perf-budget.spec.ts` pins per-keystroke timing with generous 2× budgets as
a regression scenario, but the dedicated two-fixture, quiet-machine budget gate
described in the plan is unbuilt.

### Non-Word-authored documents are deferred

Fidelity is defined against **desktop Microsoft Word** only. LibreOffice- and
Google-Docs-authored files are explicitly out of scope: the single LibreOffice
fixture, `probe3-lo-provenance` p1, scores **57.09%** and is tracked on a
separate "LibreOffice (deferred)" tab, excluded from the 1,154-page Word metric.

### Document features no fixture exercises

Verified by scanning all 109 fixtures with `zipfile` for the relevant XML/parts.
Absent from **every** fixture — so entirely unverified:

- **Charts** (`c:chartSpace`) — 0 fixtures. No render path exists either.
- **ActiveX controls** — 0 fixtures.
- **Ink annotations** (`inkml` / `w:contentPart`) — 0 fixtures.
- **Digital signatures** (`_xmlsignatures`) — 0 fixtures.
- **Pattern fills** (`a:pattFill`) — 0 fixtures (gradient fills appear in 33).

Barely covered (a single fixture — parity is essentially unproven):

- **SmartArt** — only `wild2-med-phase23-protocol.docx`. It renders solely via
  Word's pre-computed `dsp:drawing` cache (`packages/core/src/parse/document.ts:1271`);
  a SmartArt whose producer did not persist that cache would not render.
- **Legacy form fields** (`FORMTEXT`/`ffData`) — 2 fixtures.
- **Checkbox content controls** (`w14:checkbox`) — 1 fixture (toggling is tested,
  `e2e/checkbox-edit.spec.ts`); other content-control types (dropdowns, date
  pickers, rich-text SDT) are not exercised for editing.

Present and tested, for reference (not gaps): footnotes (39), endnotes (24),
comments (20), tracked changes (62), OMML math (11), VML watermarks/shapes (19),
DrawingML text boxes (19), gradient fills (33), themes (34), citations/
bibliography fields (16), embedded OLE objects (5).
