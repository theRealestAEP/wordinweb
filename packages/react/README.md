# WordInWeb

[**Live demo →**](https://wordinweb-parity.callerinfo.workers.dev/)

A Word/.docx viewer and editor for the web, embeddable as a single React component.

```tsx
import { DocxView } from "wordinweb";

<DocxView source="/report.docx" />;                            // render-only viewer
```

Editing is in an Alpha state and is strictly opt-in via the `editable` flag; the default is a pure viewer.

If you discover an edge case or perf issue create an issue and include the offending Word file.

Long-document performance work focuses on incremental parsing/layout and page virtualization; the current renderer remains DOM-based.

---

## Install

```bash
npm install wordinweb react react-dom
```

The package includes the `<DocxView />` component, optional `<DocxToolbar />`, parser, layout engine, and DOM renderer.

To match Word's glyph advances exactly, load the metric-compatible substitute fonts (Carlito ≈ Calibri, Caladea ≈ Cambria) in your app — see [Fonts](#fonts).

## Quick start

### View-only

```tsx
import { DocxView } from "wordinweb";

export function Preview() {
  return <DocxView source="/report.docx" zoom={1} style={{ height: "100vh" }} />;
}
```

`source` accepts a URL string, `ArrayBuffer`, `Uint8Array`, `Blob`, or `File` — so a drag-and-drop or `<input type="file">` handler can pass its bytes straight through.

### Editable (with the toolbar)

The editor exposes its commands through an imperative `api` handed to `onReady`. Pair it with the bundled `<DocxToolbar />` for a word-processor UI, or drive the `api` yourself.

```tsx
import { useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "wordinweb";

export function Editor() {
  const [api, setApi] = useState<DocxViewApi | null>(null);

  const download = (bytes: Uint8Array) => {
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "edited.docx";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      {api && <DocxToolbar api={api} mode="advanced" onSave={download} />}
      <DocxView
        source="/report.docx"
        editable
        commentAuthor="Ada Lovelace"
        onReady={setApi}
        style={{ height: "80vh" }}
      />
    </div>
  );
}
```

## `DocxView` props

| Prop | Type | Default | What it does |
| --- | --- | --- | --- |
| `source` | `ArrayBuffer \| Uint8Array \| Blob \| string` | — | The document: raw bytes, a `File`/`Blob`, or a URL to fetch. |
| `zoom` | `number` | `1` | Maximum zoom factor (`1` = 100%). |
| `fitWidth` | `boolean` | `true` | Scale wide pages down to fit the viewer without horizontal scrolling. Never enlarges beyond `zoom`. |
| `narrowWidth` | `number` | `820` | Container width in pixels where comments switch to the compact treatment. |
| `editable` | `boolean` | `false` | Enable editing (selection formatting, typing, save-back). Off = pure viewer. |
| `showComments` | `boolean` | `true` | Render review comments (range highlights + margin balloons). |
| `revisions` | `"final" \| "markup"` | `"final"` | Tracked-changes display: clean final text, or insertions underlined / deletions struck. |
| `commentAuthor` | `string` | `"You"` | Name stamped on comments, replies, and suggestions. |
| `onLoad` | `(info: { pageCount; document }) => void` | — | Fires once the document is parsed and laid out. |
| `onReady` | `(api: DocxViewApi) => void` | — | Fires when editing is wired up; the `api` is valid only while mounted. Only called when `editable`. |
| `onMissingFonts` | `(missing: MissingFont[]) => void` | — | Fires after render with any requested faces the browser can't render (silently substituted; layout may differ from Word). Empty array = all good. |
| `onError` | `(error: Error) => void` | — | Fires if parsing or rendering throws. |
| `className` | `string` | — | Class on the scroll container. |
| `style` | `React.CSSProperties` | — | Inline styles on the scroll container (set `height` here). |

## `DocxViewApi`

The object passed to `onReady`. Every command operates on the current selection or caret and re-lays out in place. Grouped by area:

**Character & paragraph formatting**
- `applyFormat(patch)` — bold, italic, underline, strike, `fontSizePt`, `fontFamily`, `color`, `highlight`, `verticalAlign` (super/subscript), or `clear`, over any selection (including partial runs).
- `getSelectionFormat()` — effective formatting of the selection (drives toolbar state).
- `clearFormatting()`, `changeCase("upper" | "lower" | "title")`.
- `setAlignment("left" | "center" | "right" | "justify")`.
- `setParagraphStyle(styleId | null)`, `listParagraphStyles()`, `getParagraphStyleId()`.
- `toggleList("bullet" | "number")`, `getListType()`.
- `adjustIndent(1 | -1)`, `setParagraphSpacing({ lineMultiple?, beforePt?, afterPt? })`.
- `setLink(url | null)`, `getLinkAt()`.

**Page layout**
- `setPageLayout(patch, scope?)` — margins, page size, orientation, columns, page borders. `scope` = `"document"` (all sections) or `"section"` (the caret's).
- `insertBreak("page" | "column" | "sectionNextPage" | "sectionContinuous")`.
- `insertBlankPage()` — inserts a complete blank page at the caret.
- `insertCoverPage({ title, subtitle?, author? })` — prepends an editable, styled cover page.
- `setLineNumbering(patch, scope?)`, `getLineNumbering()`.
- `openHeaderFooter("header" | "footer")`, `closeHeaderFooter()` — create or directly enter page header/footer editing; Escape closes it.

**Tables**
- `insertTable(rows, cols)`.
- `tableOp(op)` — insert/delete row/column, merge/split cells, cell vertical align, cell shading, delete table.

**Images**
- `insertImage(file)` — inserts a raster image or editable SVG icon inline at the caret, clamped to the column width (TIFF/WMF/EMF are decoded to something the browser can paint).
- `insertScreenshot()` — opens the browser screen/window/tab picker and inserts the captured frame as an editable PNG picture.
- `insertModel3D(file, poster?)` — packages a GLB as a native Office 3D model with an editable poster fallback.
- `insertOnlineVideo(url)` — inserts Word online-video metadata with a safe browser poster; double-click opens the validated HTTP(S) URL.
- `insertEmbeddedObject(file, filename?)` — embeds any file as a native OLE Package; double-click safely downloads the original file in the browser.

**Comments, footnotes & Insert content**
- `addComment(text)` — comment on the selection (`false` if nothing is selected).
- `addFootnote(text)` — footnote at the caret.
- `insertPageNumber("page" | "pageOfTotal")` — dynamic `PAGE` / `Page X of Y` field.
- `insertField(instruction, cachedResult?)`, `insertDateTime("date" | "time", picture?)` — live Word fields.
- `addBookmark(name)`, `listBookmarks()`, `insertCrossReference(name, "text" | "page")` — named bookmark targets and live `REF` / `PAGEREF` fields.
- `insertEquation(linear)`, `insertSymbol(symbol)` — native editable OMML equations and arbitrary Unicode symbols.
- `insertShape(preset, text?)` — floating editable DrawingML rectangles, rounded rectangles, ellipses, diamonds, and text boxes. Advanced mode also exposes Text Box as its own Insert control.
- `insertWordArt(text, preset?)` — editable DrawingML WordArt with plain, arch, wave, and chevron presets.
- `insertChart(data)`, `updateSelectedChart(data)` — native editable ChartML with its embedded workbook.
- `insertSmartArt(data)`, `updateSelectedSmartArt(data)` — native editable SmartArt data, layout, style, colors, and cached drawing parts.
- `setDrawingTool({ kind: "pen", color, width } | { kind: "eraser", size } | { kind: "lasso" } | null)`, `getDrawingTool()` — draw, erase, or lasso-select movable DrawingML ink.
- `arrangeObject(action)` — align the selected image, shape, or ink group to the page; rotate it 90°; or bring it to the front/send it to the back. Arrow keys nudge selected floating objects by 1px (`Shift` = 10px).
- `setDropCap("drop" | "margin" | null, lines?)` — apply, change, or remove a native Word drop cap on the caret paragraph.

**Suggesting mode (tracked changes)**
- `setSuggesting(on, author?)` — when on, edits record as OOXML `w:ins`/`w:del` instead of mutating text directly, and the view switches to markup.
- `isSuggesting()`, `acceptRevisionAtCaret()`, `rejectRevisionAtCaret()`.
- `revisionCount()`, `acceptAllRevisions()`, `rejectAllRevisions()`.

**Find & replace**
- `find(query, { matchCase? })` → match count (selects the first).
- `findStep(1 | -1)` → 1-based index of the newly selected match.
- `replaceCurrent(replacement)` → remaining matches. `replaceAll(query, replacement)` → count replaced.

**History, save & print**
- `undo()`, `redo()`, `canUndo()`, `canRedo()`.
- `save()` → edited `.docx` bytes (`Uint8Array`). Everything the model doesn't touch round-trips byte-for-byte.
- `print()` — browser print dialog / save-as-PDF of the rendered pages.
- `pageCount()` and `document` (the parsed `DocxDocument`).

### `DocxToolbar`

A ready-made formatting toolbar for an editable `DocxView`. Use `mode="simple"` for the basic Home editing strip or `mode="advanced"` for the full Home / Insert / Draw / Layout ribbon supported by the installed version.
Layout includes Word-style paper presets plus a custom width/height dialog; both document and current-section scope emit native page-size properties.

| Prop | Type | What it does |
| --- | --- | --- |
| `api` | `DocxViewApi \| null` | The api from `onReady`. |
| `onSave` | `(bytes: Uint8Array) => void` | Adds a Download button; receives `api.save()` bytes. |
| `mode` | `"simple" \| "advanced"` | Editing surface. Defaults to `"advanced"`; simple omits the Insert, Draw, and Layout ribbons. |
| `features` | `Partial<Record<ToolbarFeature, boolean>>` | Per-group toggles; every group defaults on. Set one `false` to hide it (e.g. `{ table: false, image: false }`). |
| `className` | `string` | Class on the toolbar root — handy as a scope for CSS-variable overrides. |
| `style` | `React.CSSProperties` | Inline overrides merged onto the toolbar root. |

## Theming

Every color the chrome paints — the toolbar, comment cards, selection, caret, page shadow — routes through a CSS custom property with the stock value as its fallback. The **default look is unchanged**; set any variable on an ancestor element (custom properties inherit) to retheme. Wrap the toolbar and viewer in one element and set the variables there, or set them on `:root` to theme globally.

| Variable | Default | Controls |
| --- | --- | --- |
| `--dxw-toolbar-bg` | `#f9fbfd` | Toolbar bar background |
| `--dxw-toolbar-fg` | `#3c4043` | Toolbar text & icons (icons use `currentColor`) |
| `--dxw-toolbar-border` | `#dadce0` | Toolbar borders & separators |
| `--dxw-toolbar-muted` | `#5f6368` | Secondary/label text |
| `--dxw-toolbar-z-index` | `100` | Toolbar and open-menu stacking level |
| `--dxw-accent` | `#1a73e8` | Primary accent (active tab, pills, focus rings) |
| `--dxw-accent-fg` | `#fff` | Text/icon on an accent fill |
| `--dxw-btn-active-bg` | `#dfe7f5` | Toggled button background |
| `--dxw-btn-hover-bg` | `#f1f3f4` | Button hover background |
| `--dxw-tab-active-bg` | `#e8f0fe` | Active ribbon-tab background |
| `--dxw-popover-bg` | `#fff` | Dropdown / popover surface |
| `--dxw-popover-shadow` | `0 4px 16px rgba(0,0,0,.15)` | Popover shadow |
| `--dxw-layout-menu-width` | `304px` | Layout option-menu width |
| `--dxw-layout-menu-max-height` | `480px` | Layout option-menu scroll limit |
| `--dxw-layout-preview-bg` | `#fff` | Paper fill in Layout option previews |
| `--dxw-select-height` | `26px` | Custom select trigger height |
| `--dxw-select-border` | `transparent` | Custom select trigger border color |
| `--dxw-select-radius` | `4px` | Custom select trigger corner radius |
| `--dxw-select-bg` | `transparent` | Custom select trigger fill |
| `--dxw-select-fg` | toolbar foreground | Custom select trigger text |
| `--dxw-select-font` | `13px system-ui` | Custom select trigger font |
| `--dxw-select-padding` | `0 6px` | Custom select trigger padding |
| `--dxw-select-menu-bg` | popover background | Custom select menu fill |
| `--dxw-select-menu-border` | toolbar border | Custom select menu border color |
| `--dxw-select-menu-radius` | `8px` | Custom select menu corner radius |
| `--dxw-select-menu-shadow` | popover shadow | Custom select menu shadow |
| `--dxw-select-menu-max-height` | `320px` | Custom select menu scroll limit |
| `--dxw-color-menu-width` | `236px` | Text and drawing color palette width |
| `--dxw-dialog-backdrop` | `rgba(32,33,36,.38)` | Editor dialog scrim |
| `--dxw-dialog-z-index` | `1000` | Editor dialog stacking level |
| `--dxw-dialog-width` | `420px` | Editor dialog width |
| `--dxw-dialog-radius` | `10px` | Editor dialog corner radius |
| `--dxw-dialog-field-bg` | popover background | Editor dialog field fill |
| `--dxw-canvas-bg` | `#e8eaed` | Scroll area behind the pages |
| `--dxw-page-bg` | `#ffffff` | Page paper |
| `--dxw-page-shadow` | `0 1px 3px …, 0 4px 14px …` | Page drop shadow |
| `--dxw-caret` | `#1a1a1a` | Text caret |
| `--dxw-selection` | `rgba(26,115,232,.28)` | Text-selection highlight |
| `--dxw-comment-hl` | `rgba(255,200,90,.38)` | Commented-range highlight |
| `--dxw-comment-hl-active` | `rgba(255,170,0,.55)` | Active commented-range highlight |
| `--dxw-comment-bg` | `#fff` | Comment card surface |
| `--dxw-comment-fg` | `#3c4043` | Comment card text |
| `--dxw-comment-border` | `#e0e0e0` | Comment card border |

### Dark toolbar example

```tsx
<div
  className="dxw-dark"
  style={{
    ["--dxw-toolbar-bg" as string]: "#1f2937",
    ["--dxw-toolbar-fg" as string]: "#e5e7eb",
    ["--dxw-toolbar-border" as string]: "#374151",
    ["--dxw-toolbar-muted" as string]: "#9ca3af",
    ["--dxw-accent" as string]: "#60a5fa",
    ["--dxw-btn-active-bg" as string]: "#374151",
    ["--dxw-btn-hover-bg" as string]: "#2b3644",
    ["--dxw-tab-active-bg" as string]: "#334155",
    ["--dxw-popover-bg" as string]: "#1f2937",
    ["--dxw-canvas-bg" as string]: "#111827",
  }}
>
  {api && <DocxToolbar api={api} onSave={download} />}
  <DocxView source="/report.docx" editable onReady={setApi} style={{ height: "80vh" }} />
</div>
```

Or in a stylesheet:

```css
.dxw-dark { --dxw-toolbar-bg: #1f2937; --dxw-toolbar-fg: #e5e7eb; /* … */ }
```

### Custom menus and dialogs

All visible toolbar choices use WordInWeb-rendered controls rather than the
browser's native select, prompt, or color-picker UI. `ToolbarMenuSelect` is
exported for host-app controls that should match the toolbar:

```tsx
<ToolbarMenuSelect
  value={zoom}
  ariaLabel="Document zoom"
  options={[
    { value: "0.75", label: "75%" },
    { value: "1", label: "100%" },
    { value: "1.25", label: "125%" },
  ]}
  onChange={setZoom}
/>
```

For selector-level styling, the stable hooks are
`.dxw-menu-select-trigger`, `.dxw-menu-select-menu`,
`.dxw-menu-select-option`, `.dxw-color-menu`, `.dxw-color-swatch`,
`.dxw-input-dialog`, `.dxw-input-dialog-field`, and the dialog action button
classes. Prefer the variables above for theme changes.

Layout controls also expose stable `.dxw-layout-ribbon`, `.dxw-layout-menu-trigger`,
`.dxw-layout-menu`, `.dxw-layout-menu-item`, and `.dxw-layout-preview` classes.
Their `data-dxw-layout-menu`, `data-dxw-layout-option`, and
`data-dxw-layout-preview` attributes identify each menu, action, and diagram
when a host needs a narrowly scoped style override.

## How it works

WordInWeb never converts the document to flowing HTML. It parses the OOXML into a typed model, runs a layout engine that breaks lines with real canvas metrics and paginates like Word, and renders each primitive as one absolutely-positioned element, so the browser does zero reflow. Editing mutates the retained XML tree and re-serializes only the parts it models, leaving everything else byte-for-byte intact.

## Fonts

Text is measured on a canvas before layout, so line breaks depend on the real font metrics. For Word parity, register the bundled OFL substitutes in your app entry:

```ts
import "@fontsource/carlito/400.css";   // Calibri metrics
import "@fontsource/carlito/400-italic.css";
import "@fontsource/carlito/700.css";
import "@fontsource/carlito/700-italic.css";
import "@fontsource/caladea/400.css";   // Cambria metrics
import "@fontsource/caladea/700.css";
```

For your own application, source your own fonts.

```css
@font-face {
  font-family: "Calibri";
  src: local("Calibri"), url("/fonts/Calibri.woff2") format("woff2");
  font-weight: 400;
  font-style: normal;
  font-display: swap;
}
```

Add a separate `@font-face` rule for every weight and style you use.

When the browser can't render a requested face, `onMissingFonts` reports it so you can warn the user that the on-screen layout may drift from Word.

## Development

```bash
npm install
npm test                 # core unit tests (parser + layout, deterministic measurer)
npm run build            # build the public package
```

## Rendering parity

Fidelity is measured against desktop Microsoft Word. Word exports each fixture
to PDF, WordInWeb renders the same file in the browser, and both are rasterized
and compared page-by-page. The test corpus and instructions are in
[wordinweb-parity](https://github.com/theRealestAEP/wordinweb-parity).

## License

MIT. See `LICENSE`.
