# DocxInWeb

A Word/.docx viewer and editor for the web, embeddable as a single React component.

```tsx
import { DocxView } from "@docxinweb/react";

<DocxView source="/report.docx" />;                            // render-only viewer
```

Editing is in an Alpha state and is strictly opt-in via the `editable` flag; the default is a pure viewer.

If you discover an edge case or perf issue create an issue and include the offending Word file.

Part of the broader roadmap is to migrate away from DOM rendering to just rendering on a canvas (or some hydrid thereof). Editing perf can suffer on really long documents.

---

## Install

```bash
npm install @docxinweb/react @docxinweb/core react react-dom
```

- `@docxinweb/core` — parser + layout engine + DOM renderer. Framework-agnostic, one runtime dependency (`fflate`).
- `@docxinweb/react` — the `<DocxView />` component and the optional `<DocxToolbar />`.

To match Word's glyph advances exactly, load the metric-compatible substitute fonts (Carlito ≈ Calibri, Caladea ≈ Cambria) in your app — see [Fonts](#fonts).

## Quick start

### View-only

```tsx
import { DocxView } from "@docxinweb/react";

export function Preview() {
  return <DocxView source="/report.docx" zoom={1} style={{ height: "100vh" }} />;
}
```

`source` accepts a URL string, `ArrayBuffer`, `Uint8Array`, `Blob`, or `File` — so a drag-and-drop or `<input type="file">` handler can pass its bytes straight through.

### Editable (with the toolbar)

The editor exposes its commands through an imperative `api` handed to `onReady`. Pair it with the bundled `<DocxToolbar />` for a word-processor UI, or drive the `api` yourself.

```tsx
import { useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "@docxinweb/react";

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
      {api && <DocxToolbar api={api} onSave={download} />}
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
| `zoom` | `number` | `1` | Zoom factor (`1` = 100%). Applied as a single `transform: scale()`; layout is unaffected. |
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
- `setLineNumbering(patch, scope?)`, `getLineNumbering()`.

**Tables**
- `insertTable(rows, cols)`.
- `tableOp(op)` — insert/delete row/column, merge/split cells, cell vertical align, cell shading, delete table.

**Images**
- `insertImage(file)` — inserts a `Blob`/`File` inline at the caret, clamped to the column width (TIFF/WMF/EMF are decoded to something the browser can paint).

**Comments, footnotes & fields**
- `addComment(text)` — comment on the selection (`false` if nothing is selected).
- `addFootnote(text)` — footnote at the caret.
- `insertPageNumber("page" | "pageOfTotal")` — dynamic `PAGE` / `Page X of Y` field.

**Suggesting mode (tracked changes)**
- `setSuggesting(on, author?)` — when on, edits record as OOXML `w:ins`/`w:del` instead of mutating text directly, and the view switches to markup.
- `isSuggesting()`, `acceptRevisionAtCaret()`, `rejectRevisionAtCaret()`.

**Find & replace**
- `find(query, { matchCase? })` → match count (selects the first).
- `findStep(1 | -1)` → 1-based index of the newly selected match.
- `replaceCurrent(replacement)` → remaining matches. `replaceAll(query, replacement)` → count replaced.

**History, save & print**
- `undo()`, `redo()`, `canUndo()`, `canRedo()`.
- `save()` → edited `.docx` bytes (`Uint8Array`). Everything the model doesn't touch round-trips byte-for-byte.
- `print()` — browser print dialog / save-as-PDF of the rendered pages.
- `pageCount()`, `closeHeaderFooter()`, and `document` (the parsed `DocxDocument`).

### `DocxToolbar`

A ready-made formatting toolbar for an editable `DocxView`. Ribbon-style tabs (Home / Insert / Layout) wired to the `api`.

| Prop | Type | What it does |
| --- | --- | --- |
| `api` | `DocxViewApi \| null` | The api from `onReady`. |
| `onSave` | `(bytes: Uint8Array) => void` | Adds a Download button; receives `api.save()` bytes. |
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
| `--dxw-accent` | `#1a73e8` | Primary accent (active tab, pills, focus rings) |
| `--dxw-accent-fg` | `#fff` | Text/icon on an accent fill |
| `--dxw-btn-active-bg` | `#dfe7f5` | Toggled button background |
| `--dxw-btn-hover-bg` | `#f1f3f4` | Button hover background |
| `--dxw-tab-active-bg` | `#e8f0fe` | Active ribbon-tab background |
| `--dxw-popover-bg` | `#fff` | Dropdown / popover surface |
| `--dxw-popover-shadow` | `0 4px 16px rgba(0,0,0,.15)` | Popover shadow |
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

## How it works

DocxInWeb never converts the document to flowing HTML. It parses the OOXML into a typed model, runs a layout engine that breaks lines with real canvas metrics and paginates like Word, and renders each primitive as one absolutely-positioned element, so the browser does zero reflow. Editing mutates the retained XML tree and re-serializes only the parts it models, leaving everything else byte-for-byte intact. 

See [`BLOG.md`](BLOG.md) for the pipeline and the parity work.

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

> **The Microsoft fonts in the separate parity demo are demo-only assets.** They
> are included to reproduce this project's experimental parity results, but
> they are not part of the DocxInWeb library and should not be copied into
> another application.
> Applications using DocxInWeb must source and license their own fonts. Without
> the same fonts, the viewer falls back to substitutes and glyphs or line breaks
> may differ from Word.

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

## Packages

| Package | What it is |
| --- | --- |
| `@docxinweb/core` | Parser + layout engine + DOM renderer. Framework-agnostic, zero deps besides `fflate`. |
| `@docxinweb/react` | `<DocxView />` + `<DocxToolbar />`. |

The demo, browser tests, fixture corpus, and Word parity references live in the
separate [wordinweb-parity](https://github.com/theRealestAEP/wordinweb-parity)
repository. That repository links this one as a Git submodule.

## Development

```bash
npm install
npm test                 # core unit tests (parser + layout, deterministic measurer)
npm run build            # build core + react
```

## Rendering parity

Fidelity is measured against desktop Microsoft Word. Word exports each fixture
to PDF, DocxInWeb renders the same file in the browser, and both are rasterized
and compared page-by-page. The test corpus and instructions are in
[wordinweb-parity](https://github.com/theRealestAEP/wordinweb-parity).

## License

PolyForm Noncommercial 1.0.0 — free for noncommercial use; commercial use requires permission. See LICENSE.
