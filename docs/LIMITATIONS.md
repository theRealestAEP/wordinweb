# Known limitations

The viewer is the mature surface: 1,154 Word pages compared page-by-page against
desktop Word, mean structural severity 0.026%, worst 3.95% (git `92077b9`). The
editor is newer and opt-in via the `editable` flag; its supported operations are
exactly the methods on `DocxViewApi`. Everything below is a gap, a shallow
toolbar picker, a known render residual, or an untested area, drawn from the code
and the certified parity run.

- [ ] **editing** endnote insertion: only footnotes can be authored; endnotes render (24 fixtures) but there is no `insertEndnote`.
- [ ] **editing** structured equation editing: matrices, n-ary / limit over-under groups, accents, and equation arrays open read-only; you edit a linear string and re-emit OMML, not type into the inline equation.
- [ ] **editing** rich / image paste: paste is plain-text only; copy and cut emit `text/plain` only, so pasting between two points drops run formatting and styles. Dragging an image file in from the OS does work.
- [ ] **editing** suggesting mode tracks text only: typed text and deletions record as `w:ins`/`w:del`, but formatting, tables, lists, styles, images, and layout mutate the document directly (no `rPrChange`/`pPrChange`/`tblPrChange`).
- [ ] **editing** image cropping: cropped images render, but the editor has no crop handles (only resize, reposition, wrap mode, alt-text, blip replacement).
- [ ] **editing** WordArt / watermark: only an existing VML watermark's text, rotation, and opacity are editable; no insert, no color/font control, no modern DrawingML WordArt.
- [ ] **editing** field insertion: only page-number fields insert; TOC, `DATE`, `REF`/cross-references render if present but cannot be inserted or updated.
- [ ] **editing** charts: no chart editing and no render path at all (a `chartSpace` has no code path).
- [ ] **editing** table gaps: no per-edge cell-border editing, table-style application, or numeric column-width entry (grip-drag only). Supported: insert/delete row+column, merge, split, cell shading, vertical align, including nested tables.
- [ ] **toolbar** page-border color is blue-only (`4472C4`); the command accepts any hex and always writes `val="single"` on all edges, so a color picker is trivial and line-style/per-edge control is the deeper follow-up.
- [ ] **toolbar** highlight: 5 colors (Word has ~15). Text color already uses a full picker.
- [ ] **toolbar** font list: a fixed 35-family probe list; no system-font enumeration.
- [ ] **toolbar** font sizes: fixed preset 8–48; no arbitrary point-size entry.
- [ ] **toolbar** margins: 3 presets (Normal / Narrow / Wide); the command accepts arbitrary inch values.
- [ ] **toolbar** page size: Letter / Legal / A4; no custom dimensions.
- [ ] **toolbar** columns: 1/2/3 equal width with a hardcoded 0.5in gutter; no custom widths or spacing.
- [ ] **toolbar** line + paragraph spacing: presets only (1 / 1.15 / 1.5 / 2, fixed 10pt add/remove); no custom point values.
- [ ] **toolbar** table cell fill: 5 swatches + none; no custom color.
- [ ] **toolbar** line numbers: count-by 5 or 10 presets only.
- [ ] **rendering** 22 Word pages score ≥ 0.5% severity out of 1,154; every other page is under 0.5%, most under 0.05%.
- [ ] **rendering** floor — thai p1 3.95%: licensed DokChampa is unshippable; Lao renders the bundled OFL Noto Sans Lao Looped (style-correct, not glyph-identical).
- [ ] **rendering** floor — indic p1 1.30%: Tamil Vijaya vs a scaled Latha (different licensed face).
- [ ] **rendering** floor — kashida p1 0.95%: Arabic line counts match Word; residual is concentrated join placement (rasterization-class).
- [ ] **rendering** floor — emoji p1 0.57%: Apple Color Emoji vs Segoe UI Emoji artwork; user-accepted, never pixel-matches.
- [ ] **rendering** floor — yiddish RTL body: pure stroke-rasterization of the same embedded font (line breaks and glyph positions are pixel-exact).
- [ ] **rendering** equations p1 0.82%: OMML raster-offset (equation baseline lead).
- [ ] **rendering** picture-watermark p6 0.93%: placement/opacity residual on one page.
- [ ] **rendering** dense math p8/p13 ~0.55%: compound inline n-ary shows ±2px row-pitch drift; a blanket ascent nudge fixes one sign and worsens the other, needs its own Word probe.
- [ ] **rendering** NIH contract p35/103/316/317/414 0.51–0.69%: sub-1% residuals on a 419-page real-world doc, not yet individually diagnosed.
- [ ] **rendering** rowsplit p1/p2 ~0.68%: table row split across a page boundary.
- [ ] **rendering** ieee-2col p4 0.79%: two-column science layout, sub-line drift.
- [ ] **rendering** eq-as-images p2/p5 ~0.57% and chem-omml p4 0.55%: equations/figures shipped as raster images carry a sub-pixel resample texture.
- [ ] **rendering** doerfp p8 0.82%, tblextreme p1 0.52%, phase23 p18 0.50% (the only SmartArt page): single-page residuals, undiagnosed.
- [ ] **rendering** missing Office fonts: layout matches Word via canvas metrics and metric-compatible substitutes (Carlito, Caladea, Noto Sans Lao Looped), but paint uses whatever the browser has; `onMissingFonts` warns. Real Office faces load dev-only from `/fonts-local/` and are not shipped.
- [ ] **rendering** antialiasing floor: Chrome and Word antialias strokes differently; on pixel-exact geometry this is the whole residual, and it is not closable in a browser.
- [ ] **untested** edit round-trip: no test verifies that a saved `.docx` re-opens faithfully in Word (the round-trip harness is unbuilt); serialization bugs would not be caught.
- [ ] **untested** perf budget gate: per-keystroke timing has a regression spec (`e2e/perf-budget.spec.ts`, 2× budgets) but the dedicated quiet-machine budget gate is unbuilt.
- [ ] **untested** non-Word documents: fidelity is defined against desktop Word only; the one LibreOffice fixture scores 57.09% on a separate deferred tab, excluded from the 1,154.
- [ ] **untested** zero-fixture features: charts (`c:chartSpace`, no render path), ActiveX controls, ink annotations, digital signatures, pattern fills.
- [ ] **untested** barely covered: SmartArt (1 fixture, renders only via Word's `dsp:drawing` cache), legacy form fields (2 fixtures), checkbox content controls (1 fixture; other content-control types unexercised for editing).
