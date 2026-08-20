# Tool-depth coverage matrix: LikeOffice vs Microsoft Word 365

Third measurement axis, after the rendering corpus and the editing-conformance ledger.
The rendering corpus measures what the engine can PAINT; the conformance ledger measures
what edits round-trip; this matrix measures which Word TOOLS a user (or agent) can
actually operate, and how deep each one goes.

- Reference surface: Word 365 desktop ribbon, cluster by cluster.
- Audited tree: `wordinweb-likeoffice` at `1d148f5` (branch `likeoffice`), plus the
  `likeoffice` desktop shell (`apps/desktop`).
- Read-only audit. This document is the only change on this branch.

## Status legend

| Status | Meaning |
|---|---|
| DEEP | Functional parity for the tool as users use it |
| CORE | Common workflows work; named sub-features missing |
| STUB | Renders / preserves on round-trip, but no editing surface |
| ABSENT | No render support claimed and no editing surface |

Evidence is `file:line` in the engine repo unless prefixed `desktop:`.
Key surfaces inventoried:

- Toolbar: `packages/react/src/toolbar.tsx` (tabs: home, insert, draw, layout, review,
  plus contextual object/table format — toolbar.tsx:4634)
- Host API: `DocxViewApi`, `packages/react/src/index.tsx:186-484`
- Core edit ops: `packages/core/src/edit/*` incl. the operation registry
  (`registry.ts`, 26 registered kinds)
- Collab wire: `packages/collab/src/intents.ts` (~70 hand-written intent kinds + registered ops)
- Agent surface: `packages/agent/src/capabilities.ts` (one capability row per intent kind;
  `compose.ts` additionally composes whole .docx files — compose.ts:322)
- Desktop menus: `desktop:src/main/menu.ts`

---

## 1. Clipboard

| Feature | Status | Evidence / gap |
|---|---|---|
| Copy / Cut | DEEP | Dual flavor text/plain + text/html with embedded WordprocessingML (`data-dxw-ooxml`), full source-XML fidelity — clipboard.ts:10-49, editor.ts:1825-1841 |
| Paste (in-app / Word-fidelity) | DEEP | Pasted OOXML fragment validated by the same validator as the collab wire — clipboard.ts:45-49; `pasteBlocks` intent (capabilities.ts:39) |
| Paste (external HTML / plain text) | CORE | HTML→blocks conversion — clipboard.ts:383. Gap: no image paste from clipboard files |
| Paste Special / paste-option chips | ABSENT | Keep-source / merge-format / picture / text-only choices missing. Desktop shell has only Electron's "Paste and Match Style" — desktop:menu.ts:87-91 |
| Format Painter | CORE | Single-use copy/paint — toolbar.tsx:4939-4958, index.tsx:406-408. Gap: no double-click persistent mode |
| Office Clipboard pane (24 items) | ABSENT | No multi-item clipboard history |

Counts: DEEP 2 · CORE 2 · STUB 0 · ABSENT 2

## 2. Font

| Feature | Status | Evidence / gap |
|---|---|---|
| Bold / Italic / Strikethrough | DEEP | `formatRun`/`formatRange` patch, tracked-change aware — capabilities.ts:151-167; toolbar.tsx:5002-5023 |
| Superscript / Subscript | DEEP | `verticalAlign` in run patch — capabilities.ts:160 |
| Underline | CORE | Single underline only — no underline styles (double, dotted, wavy…) or underline color |
| Font family / size | DEEP | Arbitrary family + size via API; toolbar dropdowns — toolbar.tsx:4961-4998 |
| Font color | DEEP | Palette + custom hex — toolbar.tsx:715, ColorMenu |
| Highlight | DEEP | Full highlight palette + none — toolbar.tsx:546-571 |
| Clear formatting | DEEP | `clear: true` run patch — index.tsx:364, capabilities.ts:166 |
| Change case | CORE | upper/lower/title — find.ts:112-122. Gap: sentence case, tOGGLE cASE |
| Small caps / All caps / double strike / hidden | STUB | Parsed and rendered (parse/properties.ts:192-197) but absent from `RunFormatPatch` — no way to set them |
| Character spacing / kerning / position | STUB | `letterSpacing` parsed + painted (parse/properties.ts:266, render/dom.ts:3015) but not editable |
| Text effects & typography (shadow, glow, ligatures, stylistic sets) | ABSENT | No text-effects surface |

Counts: DEEP 5 · CORE 2 · STUB 2 · ABSENT 1

## 3. Paragraph

| Feature | Status | Evidence / gap |
|---|---|---|
| Alignment (left/center/right/justify) | DEEP | `setAlignment` — index.tsx:342; `formatParagraph.align` |
| Line & paragraph spacing | DEEP | Line multiple, exact height, before/after — index.tsx:356; toolbar spacing menu incl. custom exact height dialog |
| Bulleted / numbered lists | CORE | Toggle + number-format/label/indent editing per level (`setListType`, `setNumberingLevel` registry ops; formats: decimal, roman, letters, ordinal, bullet — capabilities.ts:241-253) + restart/continue (`setNumberingRestart`). Gap: bullet/number galleries, define-new-bullet picker |
| Multilevel lists | CORE | Level up/down (`setListLevel`), per-ilvl patches. Gap: multilevel gallery, define-new-multilevel dialog, heading-numbering link |
| Indent | CORE | Half-inch steps (`adjustIndent` — index.tsx:354); exact indents only via style definitions (capabilities.ts:176-177). Gap: exact direct paragraph indent, first-line/hanging UI |
| Paragraph borders | CORE | Bottom-border divider with style/color/width/gap (`setDivider` — index.tsx:358, dialog.ts:495). Gap: other edges, box, Word's full Borders menu |
| Paragraph shading | ABSENT | Cell shading exists; paragraph shading has no surface |
| Tab stops | ABSENT | Tab characters insert (`insertSeparator` — capabilities.ts:26); tab-stop positions/leaders have no edit surface |
| Sort paragraphs | ABSENT | |
| Show ¶ (formatting marks) | ABSENT | |
| Pagination controls (widow/orphan, keep lines, keep with next, page break before) | ABSENT | `keepNext` settable only inside style definitions — capabilities.ts:178 |

Counts: DEEP 2 · CORE 4 · STUB 0 · ABSENT 5

## 4. Styles

| Feature | Status | Evidence / gap |
|---|---|---|
| Apply paragraph style | DEEP | Style dropdown + `setParagraphStyle` — index.tsx:344, toolbar.tsx:4880 |
| Create / modify / delete styles | DEEP | Registry ops `createStyle`/`modifyStyle`/`deleteStyle` with basedOn/next/quickStyle/uiPriority/linked — registry.ts:652-770; Styles pane — toolbar.tsx:3192 |
| Character styles | DEEP | `characterStyleId` run-patch property + gallery — capabilities.ts:161-165, toolbar.tsx:4914 |
| Table styles | CORE | Define borders-only table styles; conditional formats (w:tblStylePr) not expressible — capabilities.ts:195-206 |
| Numbering styles | CORE | `createStyle` type "numbering" with numId link — capabilities.ts:226 |
| Style sets / themes / Design-tab document formatting | ABSENT | |
| Style inspector / reveal formatting | ABSENT | |

Counts: DEEP 3 · CORE 2 · STUB 0 · ABSENT 2

## 5. Editing (Find / Replace / GoTo / Select)

| Feature | Status | Evidence / gap |
|---|---|---|
| Find | CORE | `find`/`findStep` select + step matches — index.tsx:368-370. Matching is body text only (headers/footers/footnotes/text boxes excluded), paragraph-local (no cross-paragraph match) — find.ts:7-9, 45-52. `matchCase` exists in the API but the popover exposes no toggle — toolbar.tsx:4243-4284 |
| Replace / Replace All | CORE | index.tsx:372-374, find.ts:89-109. Plain-text replacement; match keeps first-range formatting. UI has Find + Replace-all only — no replace-one or next/prev buttons |
| Advanced find (whole word, wildcards, format search, special characters) | ABSENT | `findAll` takes only `matchCase` — find.ts:46 |
| Go To (page/section/bookmark/…) | ABSENT | |
| Select All | DEEP | editor.ts:1013 |
| Select objects | CORE | Draw-tab Select tool + object selection model — toolbar.tsx:2058, editor.ts:1709 |
| Selection pane | ABSENT | |

Counts: DEEP 1 · CORE 3 · STUB 0 · ABSENT 3

## 6. Tables

| Feature | Status | Evidence / gap |
|---|---|---|
| Insert table | DEEP | Grid insert up to 50×50 incl. agent cell payloads — registry.ts:338, capabilities.ts:382-386 |
| Insert/delete rows & columns, delete table | DEEP | `tableOp` on wire + toolbar — intents.ts:222-249, toolbar.tsx:2863 |
| Merge / split cells | CORE | `mergeRight`/`mergeDown`/`splitCell` exist LOCALLY — tables.ts:39, toolbar.tsx:2876. Gap: they are absent from the collab wire's `tableOp` union (intents.ts:226-236), so shared documents cannot perform them; no arbitrary-rectangle merge of a selection |
| Cell shading / vertical alignment | DEEP | Tracked-change aware (w:tcPrChange) — intents.ts:234-235, 242-246 |
| Borders | DEEP | Per-edge incl. diagonals (tl2br/tr2bl), 11 styles, cell or table scope — registry.ts:856, capabilities.ts:388-393; custom border dialog toolbar.tsx:2714 |
| Table styles + style options | CORE | Apply style + Word's six look toggles — registry.ts:889-946. Gap: conditional-format definition, style gallery previews |
| Width / column widths / autofit / cell margins / header rows | DEEP | Registry ops setTableWidth (pt/pct/auto), setTableColumnWidth, setTableLayout, setTableCellMargins, setTableHeaderRows — registry.ts:947-1102; Properties dialog toolbar.tsx:2544 |
| Drag resize, floating tables, text wrapping | DEEP | `resizeTableRow`/`resizeTableColumn`/`moveTable`/`textWrapping` — registry.ts:383, capabilities.ts:82-83, intents.ts:236 |
| Sort | ABSENT | |
| Formula (=SUM(ABOVE)…) | ABSENT | |
| Convert text ↔ table | ABSENT | |
| Split table | ABSENT | |
| Distribute rows / columns | ABSENT | |
| Cell text direction | ABSENT | |

Counts: DEEP 6 · CORE 2 · STUB 0 · ABSENT 6

## 7. Illustrations

| Feature | Status | Evidence / gap |
|---|---|---|
| Pictures | CORE | Insert (PNG/JPEG/GIF/BMP/WebP, +SVG local — index.tsx:502-505), crop, resize, 5 wrap modes, alt text, rotate, z-order, exact position — capabilities.ts:66-84, registry setCrop. Gap: picture styles/effects/borders, corrections/color, remove background, stock/online pictures, compress |
| Shapes | CORE | 7 presets (line, verticalLine, rectangle, roundedRectangle, ellipse, diamond, textBox) with fill, outline, text, rotate, arrange — capabilities.ts:435, toolbar.tsx:1412. Gap: Word's ~160-shape gallery, connectors, edit points |
| Icons | CORE | User-supplied SVG insert — toolbar.tsx:5340 ("Insert SVG icon"). Gap: stock icon library |
| SmartArt | CORE | 4 layouts (process, cycle, hierarchy, list), node text/fill/format editing, data replace — capabilities.ts:284-296, smartart.ts. Gap: Word's ~130 layouts, add-shape/promote/demote UI |
| Charts | CORE | 4 native ChartML types (column, bar, line, pie) with data editor — capabilities.ts:269-282, toolbar.tsx:1757. Gap: remaining ~13 chart families, axes/legend/series formatting |
| Screenshot | DEEP | Capture screen/window/tab, insert as PNG — index.tsx:340, toolbar.tsx:2111 |
| 3D models | CORE | GLB insert + 3-axis rotation, poster — index.tsx:264, registry.ts:1175. Gap: stock 3D library, pan/zoom views |
| Ink / Draw | CORE | Pen, highlighter, eraser, lasso — toolbar.tsx:2058-2062. Gap: ink-to-shape, ink-to-math, ink replay |

Counts: DEEP 1 · CORE 7 · STUB 0 · ABSENT 0

## 8. Media

| Feature | Status | Evidence / gap |
|---|---|---|
| Online video | CORE | Word online-video metadata + browser-safe poster — index.tsx:266, toolbar.tsx:2014. Gap: in-document playback |

Counts: DEEP 0 · CORE 1 · STUB 0 · ABSENT 0

## 9. Links / Bookmarks / Cross-references

| Feature | Status | Evidence / gap |
|---|---|---|
| Hyperlink | CORE | Wrap/remove/read URL — index.tsx:350-352, links.ts. Gap: link to heading/bookmark ("Place in This Document"), e-mail links, ScreenTip, edit dialog |
| Bookmarks | DEEP | Add at caret/selection, wrap a range, list — index.tsx:202-204, capabilities.ts:53, 90 |
| Cross-references | CORE | Live text/page refs to bookmarks, PAGEREF/REF fields repage via updateFields — index.tsx:206, update-fields.ts:73. Gap: refs to headings, captions, footnotes, numbered items |

Counts: DEEP 1 · CORE 2 · STUB 0 · ABSENT 0

## 10. Comments

| Feature | Status | Evidence / gap |
|---|---|---|
| Add comment | DEEP | Selection-anchored with initials + @-mention UI — comments.ts:17, toolbar.tsx:978-1011 |
| Reply | DEEP | comments.ts:156, capabilities.ts:45 |
| Delete thread | DEEP | comments.ts:102, capabilities.ts:89 |
| Resolve / reopen | ABSENT | No resolved state surface |
| Edit an existing comment's text | ABSENT | |
| Previous / next comment navigation | ABSENT | |

Counts: DEEP 3 · CORE 0 · STUB 0 · ABSENT 3

## 11. Header & Footer

| Feature | Status | Evidence / gap |
|---|---|---|
| Edit header / footer in place | DEEP | Enter/exit editing mode, create-on-demand — editor.ts:1186-1212, `ensureHeaderFooter` capabilities.ts:88, index.tsx:425-427 |
| Page numbers | CORE | PAGE and "Page X of Y" fields at the caret — index.tsx:196, toolbar.tsx:5356. Gap: position gallery, number formats (i, ii / a, b), start-at, chapter numbers, remove-page-number |
| Different first page / odd & even pages | STUB | `titlePg` parsed, honored by layout, and preserved (parse/section.ts, sections.ts:284 preserve list); no toggle in API or UI |
| Header/footer distance from edge | ABSENT | |

Counts: DEEP 1 · CORE 1 · STUB 1 · ABSENT 1

## 12. Text (text box / WordArt / drop cap / date / object / quick parts)

| Feature | Status | Evidence / gap |
|---|---|---|
| Text box | CORE | `textBox` shape preset with editable text — capabilities.ts:435, toolbar.tsx:1497. Gap: gallery designs, linked text boxes, text direction |
| WordArt | CORE | 5 presets (plain, archUp, archDown, wave, chevron), edit text/color/opacity — capabilities.ts:442, 70-71. Gap: Word's transform/effects gallery |
| Drop cap | DEEP | drop / in-margin / none + lines, native w:dropCap — index.tsx:362, paragraph.ts |
| Date & time | DEEP | Live DATE/TIME fields with Word date pictures, refreshed by updateFields — index.tsx:200, toolbar.tsx:5388 |
| Object (embed file) | CORE | Arbitrary file as native OLE Package — index.tsx:268, toolbar.tsx:5429. Gap: create-new-from-application, display-as-icon options, Text from File |
| Cover page | CORE | Editable title/subtitle/author cover — index.tsx:418, sections.ts:44. Gap: design gallery |
| Quick Parts / Building Blocks / AutoText | ABSENT | Document-property fields (AUTHOR, TITLE, SUBJECT…) are API-insertable — fields.ts:83 — but there is no gallery or building-block store |
| Signature line | ABSENT | |

Counts: DEEP 2 · CORE 4 · STUB 0 · ABSENT 2

## 13. Symbols / Equations

| Feature | Status | Evidence / gap |
|---|---|---|
| Symbol | CORE | 24-symbol palette + arbitrary Unicode input through the undo/suggest-aware path — toolbar.tsx:1244, index.tsx:238. Gap: character-map browser, recently-used, shortcut keys |
| Equation | CORE | Native OMML: insert/edit/move/delete via linear math syntax — math.ts:181-628, capabilities.ts:43, 85-87. Word-metric math layout (layout/math.ts). Gap: built-in equation gallery, structure menus, ink equation; structures the linearizer cannot round-trip are edit-locked — editor.ts:5237 |

Counts: DEEP 0 · CORE 2 · STUB 0 · ABSENT 0

## 14. Page Setup

| Feature | Status | Evidence / gap |
|---|---|---|
| Margins | DEEP | Presets + custom + mirror margins, document or section scope — index.tsx:410, capabilities.ts:299-305, toolbar.tsx:3765 |
| Orientation | DEEP | toolbar.tsx:4131, capabilities.ts:307 |
| Paper size | DEEP | Presets + custom — toolbar.tsx:3885, capabilities.ts:306 |
| Columns | CORE | 1/2/2+divider/3 in UI, up to 12 equal via API — toolbar.tsx:4142, capabilities.ts:308. Gap: unequal widths, left/right presets, width/spacing dialog |
| Breaks | CORE | Page, column, section next-page, section continuous — index.tsx:414, sections.ts:234. Gap: even-page/odd-page section breaks, text-wrapping break |
| Line numbers | CORE | Enable, count-by, restart mode, start-at, per-section or document — sections.ts:344, toolbar.tsx:4155, capabilities.ts:322-327. Gap: suppress-for-paragraph, distance-from-text |
| Hyphenation | ABSENT | No `w:autoHyphenation` write path anywhere in edit/. Soft hyphens (U+00AD) render correctly — parse/document.ts:769-775 |
| Vertical page alignment | ABSENT | |
| Gutter | ABSENT | |

Counts: DEEP 3 · CORE 3 · STUB 0 · ABSENT 3

## 15. Page Background

| Feature | Status | Evidence / gap |
|---|---|---|
| Watermark | CORE | Text watermark on every page: diagonal/horizontal, color, opacity, remove — watermark.ts:168-177, registry.ts:1293-1350, toolbar.tsx:1540. Gap: picture watermark, font/size choices |
| Page color | ABSENT | |
| Page borders | CORE | Box border with width/color, offset from text or page, per section or document — capabilities.ts:310-319, toolbar.tsx:4004. Gap: art borders, per-edge control, styles beyond single line, shadow/3D |

Counts: DEEP 0 · CORE 2 · STUB 0 · ABSENT 1

## 16. References

| Feature | Status | Evidence / gap |
|---|---|---|
| Table of contents | CORE | Native TOC field insert with level range + leader, rebuild from headings, repage via updateFields — toc.ts, index.tsx:227-234, registry.ts:482. Gap: gallery formats, manual TOC, per-style level mapping |
| Footnotes | CORE | Insert, edit in place (caret edits mark footnotes.xml dirty — editor.ts:8052), double-click reference jumps — notes.ts:20, editor.ts:2819. Gap: next-footnote navigation, options dialog (number format, restart, location), footnote↔endnote convert |
| Endnotes | CORE | Same machinery — notes.ts:28, registry.ts:1103 |
| Insert citation | CORE | Cite a source the package already holds, display text computed, refreshed by updateFields — registry.ts:1234-1253, citations.ts:129. Gap: works ONLY against a pre-existing sources part |
| Manage sources (create/edit sources) | ABSENT | No write path for the sources part; `documentBibliography` is read-only — citations.ts:87 |
| Bibliography | STUB | Renders and updates an existing bibliography (layout/engine.ts:2595, update-fields.ts:215); no insert-bibliography tool |
| Citation style (APA / MLA / …) | ABSENT | |
| Captions | STUB | Existing "Figure N" SEQ/REF captions renumber on updateFields — update-fields.ts:73; no Insert Caption tool |
| Table of figures | STUB | TableofFigures style renders — parse/document.ts:253-255; no insert/build tool (a raw TOC field can be inserted — fields.ts:83 — but its content is never computed) |
| Index (XE / INDEX) | ABSENT | INDEX keyword passes the insert filter (fields.ts:83) but nothing ever builds it; no mark-entry |
| Table of authorities | ABSENT | |
| Researcher / Smart Lookup | ABSENT | |

Counts: DEEP 0 · CORE 4 · STUB 3 · ABSENT 5

## 17. Mailings

| Feature | Status | Evidence / gap |
|---|---|---|
| Insert merge field | CORE | MERGEFIELD with «name» placeholder, validated names — fields.ts:165-184, registry.ts:1207 |
| Data source / recipient list | ABSENT | Deliberate: the editor never writes `w:mailMerge` (settings.xml connection) — fields.ts:77-80 |
| Address block / greeting line / rules | ABSENT | |
| Preview results | ABSENT | |
| Finish & merge | ABSENT | |
| Envelopes / labels | ABSENT | |

Counts: DEEP 0 · CORE 1 · STUB 0 · ABSENT 5

## 18. Review

| Feature | Status | Evidence / gap |
|---|---|---|
| Spelling & grammar | ABSENT | The engine paints its own DOM; no proofing pass, and the hidden IME element sets `spellcheck=false` — editor.ts:593 |
| Thesaurus | ABSENT | |
| Word count | ABSENT | NUMWORDS is field-insertable (fields.ts:83) but the update pass never recomputes it — update-fields.ts:73-77 |
| Read aloud / accessibility checker | ABSENT | |
| Translate / language | ABSENT | |
| Comments | DEEP | See section 10; also on the Review tab — toolbar.tsx:4292 |
| Track changes (record) | DEEP | Suggesting mode records w:ins/w:del, formatting rPrChange/pPrChange, and table property changes (tcPrChange/tblPrChange) with per-author stamps — editor.ts:6674, suggest.ts, intents.ts:242-247 |
| Accept / Reject | DEEP | At caret + all, live revision count — editor.ts:6761-6828, toolbar.tsx:4333-4356 |
| Display for review | CORE | "final" and "markup" views — index.tsx:647-648, 1271. Gap: Original view, Simple Markup balloons, show-markup filters (by author/type), reviewing pane |
| Lock tracking | ABSENT | |
| Compare / Combine | CORE | `compareDocuments(original, revised, {author, date})` — edit/compare/. Histogram diff over a text/style/numbering fingerprint, Dice-over-bigrams inside the gaps, word-level inner diff; output is ordinary w:ins/w:del/*PrChange, so accept-all yields `revised` and reject-all yields `original` (core/test/compare-roundtrip). Text, paragraph split/merge, formatting and table rows/cells. Gap: no move detection (needs w:moveFrom/w:moveTo primitives), no Combine, headers/footers/notes/styles not compared, a paragraph holding a hyperlink or content control is struck and reinserted whole |
| Protect / restrict editing | STUB | `writeProtection` and doc-protection settings survive round-trip in the settings preserve list — docx.ts:1555; nothing enforces or edits them |

Counts: DEEP 3 · CORE 1 · STUB 1 · ABSENT 7

## 19. View modes

| Feature | Status | Evidence / gap |
|---|---|---|
| Print layout | DEEP | The whole engine: paginated, Word-metric layout with virtualized pages — layout/engine.ts |
| Editing vs Viewing mode | DEEP | `editable` prop toggles live — index.tsx:848 |
| Print / PDF | DEEP | `print()`, `exportPrintHtml()` for host PDF; desktop Export as PDF + Print — index.tsx:431-437, desktop:menu.ts:67-68 |
| Zoom | CORE | `zoom` prop + fit-width cap — index.tsx:533-547, 746. Gap: no zoom UI in the toolbar or desktop shell |
| Read mode / Web layout / Outline / Draft | ABSENT | Single view mode |
| Navigation pane (headings / pages / results) | ABSENT | |
| Ruler / gridlines | ABSENT | |
| Split / New window / View side by side | ABSENT | Desktop opens separate documents per window only — desktop:menu.ts:50 |
| Immersive / Focus | ABSENT | Desktop has OS fullscreen only — desktop:menu.ts:110 |

Counts: DEEP 3 · CORE 1 · STUB 0 · ABSENT 5

---

## (a) Status counts per section

| Section | DEEP | CORE | STUB | ABSENT |
|---|---|---|---|---|
| 1. Clipboard | 2 | 2 | 0 | 2 |
| 2. Font | 5 | 2 | 2 | 1 |
| 3. Paragraph | 2 | 4 | 0 | 5 |
| 4. Styles | 3 | 2 | 0 | 2 |
| 5. Editing | 1 | 3 | 0 | 3 |
| 6. Tables | 6 | 2 | 0 | 6 |
| 7. Illustrations | 1 | 7 | 0 | 0 |
| 8. Media | 0 | 1 | 0 | 0 |
| 9. Links/Bookmarks/Cross-refs | 1 | 2 | 0 | 0 |
| 10. Comments | 3 | 0 | 0 | 3 |
| 11. Header & Footer | 1 | 1 | 1 | 1 |
| 12. Text | 2 | 4 | 0 | 2 |
| 13. Symbols/Equations | 0 | 2 | 0 | 0 |
| 14. Page Setup | 3 | 3 | 0 | 3 |
| 15. Page Background | 0 | 2 | 0 | 1 |
| 16. References | 0 | 4 | 3 | 5 |
| 17. Mailings | 0 | 1 | 0 | 5 |
| 18. Review | 3 | 1 | 1 | 7 |
| 19. View | 3 | 1 | 0 | 5 |
| **Total** | **36** | **44** | **7** | **51** |

## (b) Ranked build list — highest-user-value gaps

Ranked by how often real documents and real sessions hit the gap, not by section order.

1. **Proofing (spell check squiggles).** The single most-used background feature in any
   word processor; wholly absent. Even a browser-dictionary pass over the layout runs
   would close most of the perceived gap.
2. **Find & Replace depth.** Expose match-case (the API already takes it — index.tsx:368),
   add next/prev buttons wired to the existing `findStep`, whole-word, search in
   headers/footers/footnotes/text boxes, cross-paragraph matches, Go To.
3. **Header/footer and page-number depth.** Different first page + odd/even toggles
   (layout already honors them — STUB), page-number formats (i, ii / start-at), position
   gallery. Nearly every long real document uses at least one of these.
4. **Tab stops + full paragraph borders/shading.** Tab stops especially: résumés, forms,
   legal documents. Today only the tab character and a bottom-border divider exist.
5. **Comment resolve / edit / navigation.** Collaborative review basics; the comment data
   model is already deep (threads, mentions, wire ops).
6. **Word count.** Trivially cheap (the layout already counts everything), constantly asked for.
7. **Table depth: merge/split on the collab wire, convert text↔table, sort.** Merge/split
   working locally but silently unavailable in shared documents (tables.ts:39 vs
   intents.ts:226) is a conformance trap as much as a feature gap.
8. **Multilevel/heading numbering galleries.** The numbering machinery is deep
   (per-level patches, restart); the missing part is the gallery UX and heading-link,
   which contracts and specs rely on.
9. **Cross-references to headings/captions + Insert Caption + table of figures.** Report
   and thesis workflows; SEQ/REF plumbing already updates existing captions.
10. **Citations pipeline completion.** Source manager (write the sources part), insert
    bibliography, style choice. Insert-citation and bibliography rendering already work;
    the missing third makes the whole cluster unusable for a fresh document.

Runners-up: paste-option chips, automatic hyphenation, underline styles + small caps,
footnote options dialog, image paste from clipboard, zoom UI, navigation pane.

## (c) ABSENT and arguably out-of-scope for v1

Separated so scope decisions stay visible. Each is ABSENT today AND has a defensible
reason to stay out of a v1 editor:

- **Combine (merge revisions from several reviewers)** — Compare itself now
  ships (see the row above); combining N documents into one revision set is
  the part still out.
- **Macros / VBA, add-ins** — code execution; out of scope by construction.
- **Mail-merge execution pipeline** (data sources, preview, finish, envelopes, labels) —
  the engine deliberately never writes `w:mailMerge` (fields.ts:77-80); MERGEFIELD
  authoring is kept, execution belongs to a host.
- **Index (XE marking + INDEX build) and Table of Authorities** — legal/back-matter
  niches with heavy dialog surface.
- **Restrict editing / IRM / digital signatures / signature line** — enforcement and
  crypto belong to a host shell; settings are round-trip preserved.
- **Thesaurus, Translate, Read Aloud, Dictate, Researcher, Smart Lookup** — online
  services, not document-engine features; the desktop shell's AI panel
  (desktop:menu.ts:98) is the LikeOffice-native answer to this cluster.
- **Master documents / subdocuments** — legacy, widely discouraged even in Word.
- **Outline / Draft / Web-layout views** — alternate projections of the model; print
  layout is the product's identity.
- **Dangerous field types** (INCLUDETEXT, INCLUDEPICTURE, DDE, MACROBUTTON, LINK) —
  excluded as an injection surface, documented posture at fields.ts:62-72; arriving
  documents still render their cached results.
- **Office Clipboard pane** — Word-ism with low modern usage.
- **Stock media libraries** (stock images, icon library, 3D library) — licensing/CDN
  concerns; the file-upload mechanisms already exist.

---

## Wave 1 delta (2026-08-08, engine 9016ffe / app bda9993)

Rows moved since the audit at ff5b00e:

- §5 Editing: Find CORE → DEEP (all stories, cross-paragraph, whole-word, matchCase UI); Go To ABSENT → CORE (page + bookmark); Replace gains options + per-story counts. Advanced-find residue: wildcards, format search, special characters. NEW FILED: replaceAll is local-only in collab (#112).
- §6 Tables: Merge/split cells CORE → DEEP (collab wire closed, e16); Sort ABSENT → DEEP (single-key); Convert text↔table ABSENT → CORE.
- §10 Comments: Resolve/reopen, edit text, prev/next navigation ABSENT → DEEP (w15:done extension marker).
- §11 Header & Footer: different-first/odd-even STUB → DEEP; page-number formats + start-at close (position gallery and chapter numbers remain).
- §18 Review: Word count ABSENT → DEEP (API + NUMWORDS/NUMCHARS recompute + app status pill); Proofing ABSENT → CORE (app-side hunspell spellcheck, native suggestion menu through the editing path, custom dictionary, language setting — grammar absent).

Revised totals: 44 DEEP / 43 CORE / 5 STUB / 46 ABSENT.

## Wave 2 delta, lane B: the citations cluster (2026-08-08, branch depth2-citations)

Build-list item 10 ("citations pipeline completion — the missing third makes
the whole cluster unusable for a fresh document") is closed: a fresh document
can now create sources, cite them, pick a style, and hold a generated
bibliography. §16 rows moved:

- Manage sources ABSENT → CORE. `createCitationSource` / `editCitationSource`
  / `deleteCitationSource` registry ops write the b:Sources custom XML part
  (ECMA-376 §22.6) — creating the whole part stack (itemN.xml + itemProps +
  rels + content types, deterministic itemID) when the package has none —
  edit/sources.ts, docx.ts sourcesTree. Delete honestly refuses while a
  CITATION still cites the tag. UI: Citations popover on the Insert tab
  (source list + New Source form — toolbar.tsx CitationsMenu). Gap: four
  source types (book, article, website, report) and the common fields; Word's
  Source Manager offers ~17 types, master/current lists, and more fields.
- Insert citation stays CORE, gap changed. The old gap (pre-existing sources
  part required) is gone; the remaining gap is the simple-tier display rule
  (APA-shaped author-date + MLA author-page; other styles fall back to the
  APA shape) — citations.ts.
- Bibliography STUB → CORE. `insertBibliography` writes Word's
  multi-paragraph BIBLIOGRAPHY field with entries GENERATED from the sources
  part (alphabetical, honest APA/MLA entry shapes — edit/bibliography.ts);
  `refreshBibliography` regenerates on updateFields, replicated (entries are
  a pure function of the sequenced sources part, so unlike TOC rebuilds it
  works in a room). Arriving Word bibliographies also render their entry
  paragraphs verbatim now (parse live-field fix). Gap: no SDT/docPartObj
  gallery wrapper or heading paragraph; entry fidelity is the simple tier.
- Citation style ABSENT → CORE. `setCitationStyle` flips
  b:Sources/@StyleName (+ SelectedStyle), one attribute both the
  parenthetical and the bibliography formatter read; APA/MLA in the UI. Gap:
  Word ships an XSL per style (Chicago, IEEE, …).

§16 counts: DEEP 0 · CORE 7 · STUB 2 · ABSENT 3. Revised totals:
44 DEEP / 46 CORE / 4 STUB / 44 ABSENT.

## Wave 2 delta, lane A (2026-08-08, branch depth2-para-refs)

Rows moved:

- §5 Editing: #112 CLOSED — replaceAll (and replaceCurrent) in collab compile
  to per-match deleteText/insertText intents (strike-then-insert while
  suggesting), per-story counts preserved; the local-only fork is gone.
  NEW: `DocxViewApi.selectRange` (#111) — exact-range selection by the
  stable-addressed wire shape, local mounts enable ids on demand (the desktop
  spellcheck's select-and-replace seam).
- §3 Paragraph: Tab stops ABSENT → DEEP (layout already rendered w:tabs incl.
  leaders and bar tabs — VERIFIED; the gap was purely the edit surface:
  registry op setTabStops, get/set API, toolbar Tabs popover). Paragraph
  borders CORE → DEEP and Paragraph shading ABSENT → DEEP (layout already
  painted pBdr all edges + w:shd — VERIFIED; registry op setParagraphBorders,
  per-edge patch + fill, Home-tab Borders menu + dialog, table border-picker
  vocabulary). Multilevel galleries: CORE gap narrowed (preset gallery
  1.1.1 / outline / Article-Section / Chapter over the existing per-level
  ops; define-new dialog and heading-link remain).
- §9 Cross-references: CORE → DEEP — targets now include headings, captions,
  and numbered items (hidden `_Ref` bookmarks via the new registered
  ensureRefBookmark op; REF/PAGEREF resolution already existed).
- §16 References: Captions STUB → DEEP (insertCaption registry op: SEQ-field
  label+number paragraph below/above, table-hoisting, Caption style
  injected, deterministic seed + updateFields renumber). Table of figures
  STUB → CORE (insertToc captionLabel → `TOC \h \z \c "Label"`, entries from
  captions, TableofFigures style, rebuild keeps the \c label; gallery
  formats and per-label UI beyond Figure remain).

Wire: ENGINE_VERSION e16 → e17 (setTabStops, setParagraphBorders,
insertCaption, ensureRefBookmark, insertToc.captionLabel).

## Wave 3 delta, lane B: DrawingML depth (2026-08-08, branch wave3-drawing)

Rows moved / gaps narrowed:

- §7 Shapes stays CORE, gap rewritten. The "~160-shape gallery" gap is
  closed: a DrawingML guide-formula evaluator (the 17 ops of §20.1.9.11 +
  pathLst→SVG incl. arcTo) over the canonical presetShapeDefinitions table
  ships 165 evaluable presets (preset-geometry.ts, preset-definitions.ts —
  generated by scripts/gen-preset-definitions.mjs; skipped: 12 actionButton*
  and the 9 bent/curved/straight connector presets, which render as lines).
  Word-authored files with ANY table preset now render the true outline —
  hand-painted geometry for 7 presets is gone — and shape text wraps in each
  geometry's own text rectangle. Authoring: insertShape accepts every table
  preset; the toolbar Shapes popover is a categorized picker (Word's 7
  categories, 147 entries, icons drawn from the presets' own paths).
  Remaining gap: elbow/curved connectors, freeform/scribble, edit points.
- §7 Charts stays CORE, gap narrowed. Native ChartML types 4 → 7: doughnut,
  area, scatter (markers+lines) plus stacked / percent-stacked bar, column
  and area, authored AND rendered (the renderer/parser already spoke them;
  authoring now writes c:grouping, stacked c:overlap=100 and the 0%
  percent-axis format). Remaining: 3-D/radar/bubble/stock/surface/ofPie/
  combo families; axes/legend/series formatting.
  CALIBRATION FILINGS (declared, not guess-calibrated): doughnut hole uses
  Word's authored default holeSize=75; stacked gapWidth stays 150 and
  overlap 100 as Word writes them, but no probe PDF has re-measured stacked
  or doughnut plots; area-series fill and scatter defaults reuse the
  line-probe-measured 3pt stroke / 9pt marker rules (probe-charts-basic),
  which the type-general comments claim but a scatter/area probe has not
  independently confirmed.
- §12 WordArt stays CORE, gap narrowed. Transforms 5 → 8 (adds circle,
  button, chevronDown — warps the painter already spoke); gallery styles
  (fill / outline / shadow combos) author native w14:textFill /
  w14:textOutline / w14:shadow run effects with the legacy w:color fallback.
  Remaining: the rest of the ~40-transform gallery, glow/reflection/3-D.
- §2 Font: Text effects ABSENT → STUB. Arriving w14 run effects now RENDER
  on any run (solid textFill wins over the fallback color; outline strokes;
  shadow paints) — but the only edit surface is WordArt insertion, not
  arbitrary runs.

Wire: ENGINE_VERSION e17 → e18 (insertShape preset vocabulary = the preset
table; insertChart/setChartData types + grouping; insertWordArt presets +
style).

## Wave 3 delta, lane A (2026-08-08, branch wave3-fields)

- §6 Tables: Formula (=SUM(ABOVE)…) ABSENT → CORE. `insertTableFormula`
  registry op writes the `=FORMULA` field (§17.16.5.22) as a w:fldSimple in
  the caret's cell, cached result EVALUATED from the containing table's cell
  texts — locale-free numeric parse (the sortTableRows rule), so every collab
  replica derives identical bytes with nothing carried; updateFields
  recomputes it (edit/formula.ts). Grammar: + - * / ^, parentheses, cell refs
  A1 / ranges A1:B3, ABOVE/BELOW/LEFT/RIGHT, SUM/AVERAGE/COUNT/MAX/MIN/
  PRODUCT, `\#` numeric picture (sections pos;neg;zero, tokens # 0 . , % and
  literals). Directional scans stop at the first empty cell (Word's
  documented SUM(ABOVE) behavior); text cells count as 0. UI: Table Format
  tab "Formula" dialog (instruction + number format, live grammar
  validation). Gap (documented at edit/formula.ts): comparison operators and
  boolean/rounding functions (IF/AND/OR/ABS/INT/MOD/ROUND…), bookmark
  operands, cross-table references, gridSpan-aware cell addressing; arriving
  fields holding those render their cached result and keep it.
- §5 Editing: Advanced find residue closed to CORE — wildcards + special
  characters (format search remains). Wildcard mode (`FindOptions.wildcards`,
  "Wildcards" checkbox) implements Word's documented subset: `? * [abc]
  [!abc] [a-z] @ < > \x ^13 ^9` — translated to a JS regex by a per-token
  compiler (never string splicing), pattern capped at 256 chars and 8
  quantifiers, malformed patterns report zero matches; always case-sensitive
  and wholeWord-free, Word's own dialog rule (the checkboxes grey out). NOT
  modeled: `{n,m}` counts, `(…)` groups/backrefs, `^0nnn` codes beyond
  ^13/^9. Literal mode now interprets Word's caret escapes: ^p (paragraph
  mark over the existing "\n" join), ^t (real w:tab — tabs and line breaks
  now join the searchable text as unaddressable characters), ^l (w:br), ^#
  ^$ ^? ^w ^s ^~ ^- ^^; unknown escapes stay literal. Limits documented in
  find.ts: a match must cover at least one real text character, and matched
  paragraph marks/tabs survive a replacement (only text is replaced);
  replacement strings take no escapes. Works in rooms unchanged — replace
  compiles to per-match intents on the originator, so no wire change.
- §16 References: Index (XE / INDEX) ABSENT → CORE (previously filed under
  "arguably out-of-scope"; the generation machinery TOC/bibliography built
  made the simple tier cheap). `insertIndexEntry` writes Word's invisible XE
  complex field (§17.16.5.31) after the selection ("Mark index entry", the
  selected text as the entry, colon = one subentry level); `insertIndex`
  builds the alphabetized index as a complex INDEX field (§17.16.5.32) —
  Index1/Index2 paragraphs (built-in styles injected), locale-free sort, and
  page numbers as PAGEREF subfields over hidden `_Idx` bookmarks wrapped
  around each mark's paragraph (the TOC entry mechanism), so the build is a
  pure function of sequenced state and replicates; updateFields harvests the
  real numbers as data and `refreshIndex` rebuilds structurally (rides the
  wire like refreshBibliography; its change test blanks harvested numbers so
  an unchanged index keeps them). Parser now treats INDEX as a live
  multi-paragraph field, so arriving Word indexes render their entries
  verbatim. Limits documented in edit/index-field.ts: main + one subentry
  level only; no cross-references (\t "See …"), page ranges (\r), \c
  columns, \h letter headings, or per-entry formatting; same-page duplicate
  marks paint a duplicated number until refresh cannot dedupe them (numbers
  are placeholders at build time).
- §14 Page Setup: Hyphenation ABSENT → STUB-plus (the SETTINGS half).
  VERIFIED first (the wave's instruction): layout ignores soft hyphens
  entirely — w:softHyphen parses to a plain U+00AD character atom
  (parse/document.ts:919), which is NOT in the layout's in-word break set
  (layout/inline.ts hyphenBreaks: only - / U+2010 between alphanumerics, and
  digit-flanked U+2013 at compat ≥ 15), and no code path paints a hyphen
  glyph at a break; nothing anywhere read w:autoHyphenation. What SHIPPED:
  the settings write path — DocxDocument.setHyphenation writes
  w:autoHyphenation / w:hyphenationZone (valued, twips) / w:doNotHyphenateCaps
  at their CT_Settings schema positions (creating + registering settings.xml
  when absent), parse + refresh read them back, `setHyphenation` registered
  op (zonePt on the wire, points convention; honest no-op via the change
  itself), api get/set, and a Layout-tab Hyphenation menu (None / Automatic /
  Automatic-keep-CAPS; zone via API only). The UI and op docs state honestly
  that this engine's layout does not hyphenate — the settings govern Word's
  rendering of the file.
  NOT shipped, filed honestly: break-opportunity honoring for EXPLICIT
  w:softHyphen. What Word does (per the engine's own probe2-hyphenation
  findings): a w:softHyphen is invisible mid-line, is a break opportunity,
  and paints a hyphen glyph when the line breaks there (a raw U+00AD typed
  into w:t is the OTHER thing — always-visible, never-breaking — and parse
  already maps it to U+2011). Implementing it needs three coupled layout
  changes: (1) zero-width measurement for U+00AD inside the cumulative
  prefix-measurement scheme (canvas measureText for U+00AD is host-dependent
  — today's behavior silently embeds that indeterminacy in line breaking for
  any wild document carrying w:softHyphen, itself a latent parity/determinism
  smell); (2) a breakAfter split at each U+00AD in the atom builder; (3) the
  line packer painting a synthetic hyphen item at a soft break AND reserving
  its width in the fit walk. PARITY RISK, stated plainly for the next parity
  wave to sentinel: wild2-sci-ieee-2col (85→91% when soft-hyphen handling
  last changed) and any corpus fixture carrying w:softHyphen are calibrated
  against the current inert behavior; the change cannot be validated without
  Word-export probes, so it should land in a parity wave with
  probe-softhyphen fixtures, not here.
  assignFreshTracked + core StableIds.unassign): a refresh inside a mutation
  auto-assigns sequential ids to fresh nodes, and after an earlier intent's
  partly-consumed carried batch those autos could land inside the NEXT
  batch's carried range — reassign then threw and the intent was rejected as
  "apply failed" on every replica (any second large registered insert in one
  room could hit it, e.g. TOC after TOC). Fresh nodes' autos are now dropped
  before the carried batch lands; overflow nodes re-auto-assign above it.

## Wave 3 delta, lane C: page-number and header/footer galleries (2026-08-08, branch wave3-galleries2)

Three new registered operations (`packages/core/src/edit/hf-gallery.ts`),
each composed entirely from machinery this engine already has — no new field
instructions, no new part kinds, no ENGINE_VERSION fence (an old peer that
does not know the kind rejects it cleanly through the existing registry
gate, the same shape the citations-cluster wave shipped without a bump).

- §11 Header & Footer, Page numbers row: gap "position gallery,
  remove-page-number" CLOSES (the row's other historical gaps — formats,
  start-at — were already closed by wave 1 without the row text being
  rewritten; this doc's per-row tables are the ORIGINAL audit snapshot,
  never edited in place — see the file header. Current state is always the
  union of the base table plus every delta below it).
  `insertPageNumberPosition` (registry.ts) inserts a single live PAGE field
  into the header ("top") or footer ("bottom"), aligned left/center/right —
  Word's six "Plain Number" gallery entries — by composing `ensureHfPart`
  (the same part-creation `ensureHeaderFooter` uses) with the PAGE field
  vocabulary fields.ts already allows. A pick REPLACES the part's content,
  matching Word's own gallery (and insertWatermark's precedent), so the
  operation always applies rather than checking an unchanged-content no-op.
  `removePageNumbers` strips PAGE/NUMPAGES content from every header and
  footer part — both this engine's own `w:fldSimple` fields AND a same-
  paragraph Word-authored complex-field span (`w:fldChar` begin/instrText/
  separate/end), plus the adjacent "Page "/" of " literals the "Page X of Y"
  form writes, removed only when directly adjacent to a removed field so
  unrelated text is untouched. UI: Insert ▸ Page number gains a Top/Bottom ×
  Left/Center/Right group and a Remove entry (toolbar.tsx PageNumberMenu).
- Header & Footer preset gallery (Insert ▸ Header & footer): NEW capability,
  not tied to a specific pre-existing gap row (the original audit had no
  "preset gallery" line item for headers/footers beyond page numbers).
  `insertHeaderFooterPreset` replaces a header or footer's content with one
  of four layouts — blank, centered title, title + date, three-column —
  composed from literal placeholder text, the existing DATE and PAGE field
  vocabulary, and `setTabStops` (the wave-2 tab-stop op) for the
  three-column layout's center/right tab positions, computed from the
  document's own page width and margins. UI: HeaderFooterMenu gains a
  four-entry preset group per band.
- §12 Text, Cover page row: gap "design gallery" NARROWS. `insertCoverPage`
  (a hand-written intent, unchanged wire shape) gains an optional
  `content.layout`: "title" (unchanged, the default), "banner" (the title
  paragraph shaded into a colored band via `setParagraphBorders`' shading —
  reusing the wave-2 paragraph-shading op rather than inventing a second way
  to paint a fill), and "sidebar" (left-aligned, lower on the page, with a
  colored left accent rule via the same op's per-edge borders). Remaining
  gap: the rest of Word's ~16-design gallery; these three distinct layouts
  are the simple tier. UI: CoverPageMenu gains a three-way layout picker.

Filed, not shipped: **Quick Parts** (save-selection-as-building-block +
insert-from-gallery), item 3 of this wave's priority list. VERIFIED first,
per the wave's instruction: this engine has ZERO existing support for
OOXML's glossary document (`word/glossary/document.xml`, ECMA-376 §17.12) —
no parser read path, no part in `DocxDocument`'s tracked-part model
(`hfParts`/`footnotesRoot`/`commentsRoot`/`sourcesTree` are the precedent
for what a new part kind costs: each needed its own content-type override,
relationship, and — for anything the caret can enter —
`editableRoots()`/stable-id integration). Building the real glossary part
with `w:docPart`/`docPartPr` gallery metadata is a wave-sized lane of its
own, comparable to the wave-2 citations cluster.

The instruction's own escape hatch — "an honest document-settings-based
store is acceptable" — was evaluated concretely: OOXML's CT_Settings does
have a legitimate, schema-legal extension point for exactly this (`w:docVars`,
§17.15.1.34, arbitrary name/value string pairs Word itself exposes as VBA
`ActiveDocument.Variables`), and the natural design reuses machinery this
engine already has end to end: `saveQuickPart` would validate a selection's
serialized OOXML through the SAME `validatePastedOoxml` gate `pasteBlocks`
already puts untrusted fragments through, and `insertQuickPart` would splice
it back exactly the way `pasteBlocks`' apply does (parse, validate, splice,
assign carried ids to the fresh nodes). That reuse is sound. What blocked
shipping it in this wave is placing `w:docVars` at its correct position in
CT_Settings' long sequence — every settings writer in this codebase
(`SETTINGS_BEFORE_MIRROR`, `SETTINGS_BEFORE_EVEN_AND_ODD` in docx.ts) hand-
maintains an exact ordered predecessor list so Word never has to repair the
file, and `w:docVars` sits deep in that sequence (after `w:compat`, before
`w:rsids`) — far enough past the existing lists that transcribing it from
memory rather than a schema reference was a real correctness risk to a real
Word file, not a cosmetic one. Recommended follow-up: a wave that opens the
ECMA-376 CT_Settings schema (or a real Word-saved settings.xml with docVars
present) to get the ordering right, then wires `saveQuickPart` /
`insertQuickPart` / `deleteQuickPart` as hand-written intents on top of the
verified write path.

## Wave 4 delta, engine lane (2026-08-08, branch wave4-eq-notes)

Three rows this wave, all VERIFY-first per instruction. Wire: ENGINE_VERSION
e18 → e19 (setFootnoteOptions, setEndnoteOptions; formatRun/formatRange gain
the textEffect patch field).

- §13 Symbols/Equations: Equation's gap narrows. VERIFIED first (hand-traced
  the parser/linearizer, cross-checked against math-corpus.test.ts's
  round-trip gate): the linear grammar names fraction, superscript,
  subscript, radical, the three n-ary operators (integral/sum/product plus
  the rest of NARY_CHRS), one delimiter pair, a stacked limit, and a matrix —
  every one of those round-trips byte-for-byte. A "Structures" gallery on the
  existing EquationMenu (toolbar.tsx) inserts a ready-made template with □
  placeholders for each, through the SAME api.insertEquation path the manual
  linear-text field already used — no core changes needed; the templates are
  just the grammar's own syntax. Word's "Function" structures (sin, cos,
  log…) are deliberately NOT in the palette: they author m:func, which
  edit/math.ts does not model — isLinearSafe already refuses arriving m:func
  equations (math-corpus.test.ts: "dense refuses four m:func equations"), so
  a Function button would insert an equation the user could never edit again.
  Extending the grammar to cover m:func would touch tagOf/buildOmml/
  ommlToNodes/sameShape/withSource/graft plus new parser syntax — not the
  "small, fully tested" bar the wave's instruction set for a grammar
  extension, so it stays filed as a gap rather than rushed.
- §16 References: Footnotes/Endnotes gap narrows ("options dialog (number
  format, restart, location)"). VERIFIED first: numFmt/numStart already drove
  the painted mark (formatNoteMark reuses the page-number formatter plus a
  "chicago" symbol style — parse/section.ts already read both into
  SectionProps), but numRestart and pos were parsed nowhere, and none of the
  four had a write path — footnotePr/endnotePr sat in the sectPr
  round-trip-preserved-verbatim tree with no editor sitting on top. Shipped:
  registered ops setFootnoteOptions / setEndnoteOptions write w:footnotePr /
  w:endnotePr at their CT_FtnProps/CT_EdnProps schema position (pos, numFmt,
  numStart, numRestart — ECMA-376 §17.11.17/19/21), document- or
  section-scoped like setPageNumberFormat; a "Note options" popover beside
  the existing footnote/endnote insert controls exposes all four fields for
  both note types. assignNoteNumbers (layout/engine.ts) now honors
  numRestart="eachSect" (resets the running mark counter to that section's
  own numStart when a section begins) — a section-scoped counter reset was
  cheap since sections are already visited in document order. numRestart=
  "eachPage" round-trips but is NOT laid out, filed honestly in both the code
  and the op's doc comment: mark numbers are assigned in one whole-document
  pass before pagination runs, so which page a note will land on isn't known
  yet: honoring it needs numbering folded into the pagination pass itself, a
  materially bigger change than a settings write path. pos (page-bottom vs.
  beneath-text; section-end vs. document-end) round-trips without changing
  layout — this engine always places footnotes at the page bottom and
  endnotes at the document end, regardless of the value written.
- §2 Font: Text effects STUB → CORE. The wave3-drawing STUB note said "the
  only edit surface is WordArt insertion, not arbitrary runs." VERIFIED:
  formatRange's rPr patch was the right fit, not a new registry op — applying
  an effect to PART of a run splits the run (the registry can't do that;
  formatRange already carries before/middle/after piece ids for exactly this
  reason, the same argument characterStyleId's doc comment already makes).
  RunFormatPatch gains textEffect (outline + shadow, null clears); a new
  edit/drawings.ts function applyRunTextEffect writes/strips the w14
  elements directly on an ordinary run's rPr, careful to distinguish w14:
  shadow from the schema-unrelated legacy w:shadow toggle that shares its
  local name (§17.3.2.36 — a real collision the code comments and a
  dedicated test call out). A small Home-tab "Text Effects and Typography"
  preset row (toolbar.tsx) offers outline, shadow, an outline+shadow combo,
  a bold-outline combo, and one fixed-palette white/blue/shadow combo — six
  buttons including a "no effect" clear, not WordArt's ~20-swatch gallery.
  collab/validate.ts bounds the new nested patch shape (outline color/width)
  the same way it already bounds characterStyleId. Deliberately NOT
  refactored: WordArt insertion's own applyWordArtStyle (edit/drawings.ts)
  duplicates the outline/shadow XML construction rather than sharing it —
  only the second occurrence of that shape, not worth abstracting yet, and
  touching the existing (tested, working) WordArt path was unnecessary risk
  for this change.

## Wave 5: Quick Parts / Building Blocks (2026-08-08, branch quick-parts)

Closes the item wave3-galleries2 filed rather than shipped: "this engine has
ZERO existing support for OOXML's glossary document." §12 row moves:

- Quick Parts / Building Blocks / AutoText ABSENT → CORE. The real OOXML
  glossary part (`word/glossary/document.xml`, ECMA-376 §17.12), not the
  filed `w:docVars` fallback — DocxDocument.glossaryTree creates the whole
  part stack (relationship + content-type override, the createNotesPart
  discipline; no datastore-item companions, no `w:docPartPr/w:guid`, unlike
  the citations-cluster sources part) when the package has none, and retains
  an arriving one byte-stably (edit/quick-parts.ts, docx.ts glossaryTree).
  `createBuildingBlock` validates the current selection's serialized OOXML
  through `decodeClipboardOoxml` — the SAME allowlist gate `pasteBlocks`/
  paste already put untrusted fragments through — before it enters the
  part, and is refused for a name already taken (the tag-collision
  predicate `createCitationSource` established). `insertBuildingBlock`
  deep-clones a named docPart's stored blocks after the caret's paragraph
  (the `insertBibliography`/`insertCoverPage` precedent); its content is a
  pure function of already-synced glossary state, so — unlike `pasteBlocks`
  — it carries no OOXML of its own over the wire, only the standard
  carried-id BUDGET (`blockCount`, the `entryCount` pattern). All three ride
  the generic registered-operation machinery (collab validate/apply/
  transform, agent capability rows) with zero per-kind wiring beyond the
  registry declaration. `deleteBuildingBlock` is an honest no-op when the
  name is absent. UI: Insert-tab "Quick Parts" menu (toolbar.tsx
  QuickPartsMenu) — save-selection-as-Quick-Part form (name + category,
  category default "General"), a gallery grouped by category with
  Insert/Delete affordances, the CitationsMenu popover idiom. Word interop:
  an arriving Word glossary part's companion parts (its own styles/
  settings/fontTable/rels, which real Word writes and this engine never
  reads) survive save() byte-stably through the generic untouched-part
  path, not any glossary-specific code. Gap: no `w:docPartPr/w:guid` or
  gallery vocabulary beyond the one Word uses for a user's own "Save
  Selection to Quick Part Gallery" (`docParts`); Word's built-in Building
  Blocks Organizer content (page-number/watermark/cover-page galleries this
  engine already authors directly) is not imported as glossary entries; no
  in-place editing of a stored block short of delete-and-resave.

§12 counts: DEEP 2 · CORE 5 · STUB 0 · ABSENT 1.

---

## Scope-out costing (2026-08-12) — corrections to this document

The "deliberate scope-outs" above were a hand-wave. They have now been costed in
WAVE LANES, sized against what actually shipped (index-field.ts 431 lines,
sources.ts 337 + bibliography.ts 208, hf-gallery.ts 296, quick-parts.ts 223 —
each one lane of core + toolbar + tests). Full analysis:
scratchpad/scope-out-costing.md. Seven rows here are wrong:

- **Index — ALREADY BUILT. This document is stale.** Wave 3 lane A moved it
  ABSENT → CORE. Delete the scope-out row. Only Table of Authorities remains,
  and it is now ~½-1 lane because index-field.ts is a working template for the
  same machine (mark field + build field + PAGEREF over hidden bookmarks +
  refresh on the wire); TOA adds only category grouping and the passim rule.
- **Mail-merge execution — the stated reason is factually wrong.** The posture
  in fields.ts:66-90 guards RESOURCE-NAMING instructions: w:mailMerge persists a
  data-source CONNECTION that the next opener resolves, which is the injection
  surface. Merging against a CSV the user picks in a file dialog persists no such
  element — only substituted text enters the document. Preview is ~1 lane;
  finish-to-N-documents is ~½ more on compose.ts's existing precedent.
- **Compare / combine — over-costed. BUILT 2026-08-12, and the costing held.**
  suggest.ts already exports
  insertSuggestedText, deleteSuggestedRange, markParagraphGlyph, the four
  record*Change functions, collectRevisions and acceptRevision, all carrying
  per-author RevisionMeta. That IS Compare's output format. The only missing
  piece is alignment — no diff code exists anywhere in the tree — and diff
  (BSD-3), fast-diff and diff-match-patch (Apache-2.0) are MIT-compatible. ~1
  lane for the text tier. The risk is alignment QUALITY, not plumbing.
  Outcome: one lane, `diff` taken for the LCS core only. The alignment call
  was histogram (not patience — patience finds no anchor at all among three
  identical "Introduction" headings), and the prediction that quality was the
  risk was right: the plumbing worked early, and the effort went into the
  coalescing pass that stops a rewritten sentence coming out as confetti.
- **Restrict editing / IRM / signatures — one row conflating three costs.**
  Tier A (honor + author w:documentProtection as an editor mode) is ~½ lane and
  maps onto surfaces that exist; ship it with Word's own honesty that the
  password is a hash, not encryption. Tier B (real ECMA-376 Part 2 package
  encryption) is 1-2 lanes plus a crypto review. Tier C (IRM, digital
  signatures) needs AD RMS/Azure and a certificate trust chain — permanent.
- **"Online services" conflates two different things.** Read Aloud is platform
  speech synthesis, not a cloud call — HOURS, and selectRange (Wave 2 lane A) is
  exactly what word-level highlight-as-you-speak needs. It is accessibility, not
  convenience. Thesaurus/Translate/Researcher/Smart Lookup are already ANSWERED
  by the AI panel; listing them ABSENT undercounts the product against itself.
- **Master documents — permanent, but for a better reason.** w:subDoc is a
  cross-file reference, the same class fields.ts already excludes alongside
  INCLUDETEXT/INCLUDEPICTURE/DDE/LINK. Building it would REVERSE a documented
  security posture. State that, not "legacy and discouraged".
- **VBA — permanent, and an unclaimed win.** No MIT-compatible interpreter of
  any maturity exists, and executing code from arriving documents inside Electron
  is the macro-virus vector Word spent two decades containing. BUT
  buildPackageFiles (docx.ts:2316) starts from `{ ...this.pkg.raw() }`, and
  nothing special-cases vbaProject.bin or .docm — so **macros already round-trip
  byte-identically today**. "Preserved, never executed" is a true claim that
  costs nothing and is a better product line than silence.

Legacy import (.doc/RTF): the cheap path is detecting `soffice` on PATH and
shelling out to `--convert-to docx` (MPL-2.0, separate process, no linking, so
MIT is unaffected) — one filter to widen at main/index.ts:39, and it unlocks a
dozen formats at once. A text-only .doc reader would be worse than none for a
product whose identity is 0.004% divergence.
