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
| Compare / Combine | ABSENT | |
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

- **Compare / Combine documents** — a diff engine of its own; distinct product surface.
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
