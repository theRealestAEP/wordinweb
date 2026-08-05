# WordInWeb agent interface

## Contents

- [Execution surfaces](#execution-surfaces)
- [Tool contracts](#tool-contracts)
- [Inspection types](#inspection-types)
- [Text projection](#text-projection)
- [Patching a projection](#patching-a-projection)
- [References](#references)
- [Editing](#editing)
- [Compose generic primitives](#compose-generic-primitives)
- [Complete operation catalog](#complete-operation-catalog)
- [Nested value shapes](#nested-value-shapes)
- [JavaScript API](#javascript-api)
- [Session modes](#session-modes)

## Execution surfaces

Use the `word_document_*` tools when the host exposes them. Use `AgentDocument` when JavaScript access is available. Both surfaces use the same document state, references, schemas, and revisions.

The host supplies either a blank DOCX, an existing DOCX, a local browser session, or a collaborative session.

## Tool contracts

### `word_document_capabilities`

Return edit descriptions and closed JSON Schemas.

```json
{
  "category": "text | paragraph | review | table | insert | drawing | math | document",
  "kind": "optional exact operation name"
}
```

Omit both fields to retrieve the complete catalog. Prefer `kind` for the smallest response.

### `word_document_compose`

Create the complete structure of a new `AgentDocument.create()` document in one atomic request.

```ts
{
  revision: string;
  page?: PageLayoutPatch;
  body: AgentComposeBlock[];
  header?: AgentComposeBlock[];
  footer?: AgentComposeBlock[];
  firstHeader?: AgentComposeBlock[];
  firstFooter?: AgentComposeBlock[];
}
```

Composition accepts 1–500 body blocks and up to 100 header or footer blocks.
It validates the complete closed schema before it changes the document. Register
image assets first, then use their `assetRef` values in image blocks.

```ts
type AgentComposeBlock =
  | { type: "paragraph"; text?: string; runs?: AgentComposeRun[]; styleId?: "Normal" | "Title" | "Subtitle" | "Heading1" | "Heading2" | "Heading3"; align?: "left" | "center" | "right" | "both"; spacing?: { beforePt?: number; afterPt?: number; lineMultiple?: number }; list?: "bullet" | "number" }
  | { type: "heading"; level: 1 | 2 | 3; text: string }
  | { type: "table"; rows: Array<Array<string | AgentComposeCell>>; headerRows?: number; headerFill?: string; headerTextColor?: string; columnWidths?: number[] /* pixels */ }
  | { type: "equation"; mathText: string; align?: "left" | "center" | "right" }
  | { type: "chart"; chart: ChartData; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "smartArt"; smartArt: SmartArtData; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "image"; assetRef: string; widthPx: number; heightPx: number; alt?: string; align?: "left" | "center" | "right"; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; position?: { xPx: number; yPx: number } }
  | { type: "shape"; preset: "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox"; text?: string; textStyle?: AgentComposeTextStyle; widthPx?: number; heightPx?: number; position?: { xPx: number; yPx: number }; fill?: string | null; line?: { color: string; widthPx: number; dash: "solid" | "dashed" | "dotted" } | null; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; order?: "front" | "back" }
  | { type: "wordArt"; text: string; preset: "plain" | "archUp" | "archDown" | "wave" | "chevron"; widthPx?: number; heightPx?: number; position?: { xPx: number; yPx: number }; rotation?: number; fill?: string; opacity?: number; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; order?: "front" | "back" }
  | { type: "pageNumber"; fieldKind: "page" | "pageOfTotal"; align?: "left" | "center" | "right"; color?: string; fontSizePt?: number; fontFamily?: string; bold?: boolean }
  | { type: "pageBreak" };

interface AgentComposeRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  highlight?: string;
  fontSizePt?: number;
  fontFamily?: string;
}

interface AgentComposeCell {
  text?: string;
  runs?: AgentComposeRun[];
  fill?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
}

interface AgentComposeTextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSizePt?: number;
  fontFamily?: string;
  align?: "left" | "center" | "right";
}
```

Use `heading` blocks for a native outline. Set both dimensions when a chart or
SmartArt block needs a custom size. Use `align` for inline visual components.
Use `wordArt` in the header with `wrap: "behind"` and a low `opacity` for a
faint repeating watermark. Use a page-sized `shape` with `wrap: "behind"` for
a positioned background panel. Floating positions use page coordinates. A
Letter-page panel uses `816` by `1056` pixels at `{ xPx: 0, yPx: 0 }`. Set
`line: null` for a borderless shape. Use `textStyle` for shape text. Use
`firstHeader` or `firstFooter` when a cover needs different page furniture.

The result contains the new revision, counts, a complete `overview`, up to 100
compact `createdObjects`, and `createdObjectsTruncated`. Use these summaries
before another overview or story read.

```ts
interface AgentComposeResult {
  revision: string;
  status: "applied";
  blocks: number;
  components: number;
  overview: AgentOverview;
  createdObjects: Array<{ ref: string; editRef?: string; type: string; label?: string }>;
  createdObjectsTruncated: boolean;
}
```

### `word_document_inspect`

Inspect document content progressively.

```json
{ "kind": "context", "maxBlocks": 100, "maxCharacters": 24000 }
{ "kind": "context", "stories": ["body", "header:rId7"], "include": ["bookmarks", "objects"] }
{ "kind": "overview" }
{ "kind": "read", "story": "body", "maxBlocks": 20, "maxCharacters": 12000 }
{ "kind": "read", "story": "body", "cursor": { "value": "returned cursor" } }
{ "kind": "search", "query": "revenue", "maxResults": 50 }
{ "kind": "object", "ref": "object reference" }
{ "kind": "spatial", "pages": { "start": 1, "count": 10 }, "includeOverlaps": true }
```

Limits:

- `context.stories`: 1–100 unique story IDs
- `context.maxBlocks`: 1–200 across the complete response
- `context.maxCharacters`: 1–100,000 across the complete response
- `read.maxBlocks`: 1–200
- `read.maxCharacters`: 1–100,000
- `search.query`: 1–1,000 characters
- `search.maxResults`: 1–500
- `spatial.pages.count`: 1–100

### `word_document_edit`

Apply one transaction against an inspected revision.

```json
{
  "revision": "17",
  "operations": [
    { "kind": "insertText", "at": { "blockRef": "block:1", "runRef": "run:2", "offset": 0 }, "text": "Quarterly report" }
  ]
}
```

A transaction accepts 1–100 operations. The tool schema contains the closed schema for every operation and nested value. Runtime validation applies the same contract before mutation.

### `word_document_project`

Render one story as deterministic text plus a line anchor map. Read-only.

```json
{ "mode": "md" }
{ "mode": "text", "story": "body", "maxBlocks": 200, "maxCharacters": 24000 }
{ "mode": "md", "cursor": { "value": "returned cursor" } }
{ "mode": "outline" }
```

Limits:

- `mode`: `text`, `md`, or `outline`. Default `md`.
- `story`: a story ID from `overview`. Default `body`.
- `maxBlocks`: 1–2,000 top-level blocks per window
- `maxCharacters`: 1–200,000 per window

The result carries the projected `text`, the `revision` it belongs to, and one
`anchors` entry per line. See [Text projection](#text-projection).

### `word_document_patch`

Rewrite lines of a projection window. One transaction, all or nothing.

```json
{
  "revision": "17",
  "mode": "md",
  "edits": [{ "startLine": 4, "endLine": 4, "newText": "Adopt the managed platform." }]
}
```

```json
{
  "revision": "17",
  "mode": "md",
  "diff": "@@ -4,1 +4,2 @@\n-Adopt the platform.\n+Adopt the platform.\n+Name an owner.\n"
}
```

Send `edits` or `diff`, never both. Repeat the `story`, `mode`, and `cursor`
of the projection the lines came from. Set `suggest: true` to record the patch
as tracked changes. A patch accepts 1–100 hunks and compiles to at most 100
operations.

The result contains the new `revision`, the applied operation kinds, and the
refreshed `projection` of the same window, so a second call to re-anchor is
unnecessary. See [Patching a projection](#patching-a-projection).

### `word_document_asset`

Read a document asset returned by object inspection.

```json
{ "ref": "asset:1" }
```

The result contains `ref`, `mediaType`, and base64 bytes.

### `word_document_save`

Serialize the current document.

```json
{}
```

The result contains DOCX `mediaType` and base64 bytes.

## Inspection types

### Compact context

`context` returns all non-empty stories by default. It omits optional fields
whose values are empty. The default global budget is 100 blocks and 24,000
characters.

```ts
interface AgentContextResult {
  revision: string;
  contents: Array<{
    story: string;
    kind: "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";
    blocks: Array<
      | {
          type: "paragraph";
          ref: string;
          text: string;
          range?: { start: number; end: number; total: number };
          runs: Array<{ ref: string; start: number; end: number; wireLength?: number }>;
          styleId?: string;
          outlineLevel?: number;
          list?: { numId: number; level: number };
          bookmarks?: string[];
          objects?: Array<{ ref: string; editRef?: string; type: string; label?: string }>;
        }
      | { type: "table"; ref: string; rows: number; columns: number }
    >;
    next?: { value: string };
  }>;
  truncated: boolean;
  remainingStories?: string[];
}
```

Use a returned `next` cursor with a detailed `read` request for the same story.
Set `includeEmpty: true` only when empty paragraph targets matter.

### Overview

`overview` returns:

```ts
interface AgentOverview {
  revision: string;
  sections: number;
  mirrorMargins: boolean;
  sectionLayouts: Array<{
    index: number;
    page: { width: number; height: number };
    margins: { top: number; right: number; bottom: number; left: number };
    headerDistance: number;
    footerDistance: number;
    gutter: number;
    columns: { count: number; space: number; widths?: number[]; spaces?: number[]; separator?: boolean };
    titlePage: boolean;
    pageNumberStart?: number;
    pageNumberFormat?: string;
    breakType?: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
    verticalAlignment?: "top" | "center" | "both" | "bottom";
    pageBorders?: {
      top?: AgentBorder;
      bottom?: AgentBorder;
      left?: AgentBorder;
      right?: AgentBorder;
      offsetFrom: "text" | "page";
    };
    lineNumbering?: { countBy: number; start: number; distance: number; restart: "continuous" | "newPage" | "newSection" };
  }>;
  stories: Array<{
    id: string;
    kind: "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";
    blocks: number;
  }>;
  blocks: { paragraphs: number; tables: number };
  characters: number;
  comments: number;
  components: Record<string, number>;
  objectCounts: Record<string, number>;
  outline: Array<{ ref: string; level: number; text: string; styleId?: string }>;
}

interface AgentBorder {
  style: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave" | "dotDash" | "dotDotDash" | "thinThickSmallGap" | "triple";
  width: number;
  color: string;
  space: number;
  rawWidth?: number;
}
```

### Read result

`read` returns paragraphs and tables plus a cursor when more content remains.

```ts
interface AgentReadResult {
  revision: string;
  story: string;
  blocks: Array<AgentParagraph | AgentTable>;
  next?: { value: string };
  truncated: boolean;
}

interface AgentParagraph {
  type: "paragraph";
  ref: string;
  editable: boolean;
  story: string;
  styleId?: string;
  outlineLevel?: number;
  alignment?: string;
  list?: { numId: number; level: number } | null;
  text: string;
  textRange: { start: number; end: number; total: number };
  runs: AgentRun[];
  bookmarks: string[];
}

interface AgentRun {
  ref: string;
  editable: boolean;
  paragraphStart: number;
  paragraphEnd: number;
  wireLength: number;
  text: string;
  formatting: Record<string, unknown>;
  hyperlink?: { href?: string; anchor?: string };
  components: Array<{ ref: string; editRef?: string; type: string; label?: string }>;
}

interface AgentTable {
  type: "table";
  ref: string;
  editable: boolean;
  story: string;
  rows: number;
  columns: number;
  styleId?: string;
  floating: boolean;
  cells: Array<{ row: number; column: number; blocks: string[] }>;
}
```

Offsets use the run's Word wire length. Use the returned `wireLength` and edit schemas when text contains fields, equations, tabs, or drawings.

### Search result

```ts
interface AgentSearchResult {
  revision: string;
  matches: Array<{
    blockRef: string;
    runRef: string;
    editable: boolean;
    start: number;
    end: number;
    excerpt: string;
  }>;
  truncated: boolean;
}
```

### Object result

```ts
interface AgentObjectResult {
  revision: string;
  ref: string;
  editRef?: string;
  type: string;
  detail: Record<string, unknown>;
}
```

Objects include breaks, fields, equations, note references, ruby text, images, anchored drawings, text boxes, WordArt, charts, SmartArt, vector drawings, embedded objects, 3D models, and web videos. The detail can contain an `assetRef` or a nested story ID.

### Spatial result

```ts
interface AgentSpatialResult {
  revision: string;
  layout: { quality: "exact" | "approximate"; profile: string };
  totalPages: number;
  pages: Array<{ index: number; width: number; height: number }>;
  objects: Array<{
    ref: string;
    editRef?: string;
    type: string;
    page: number;
    bounds: { x: number; y: number; width: number; height: number };
    rotation?: number;
    layer: "behind-text" | "body" | "in-front-of-text";
    zOrder?: number;
  }>;
  overlaps: Array<{
    objects: [string, string];
    overlapBounds: { x: number; y: number; width: number; height: number };
    overlapArea: number;
    topObject?: string;
  }>;
}
```

Browser layout uses canvas metrics. Headless layout uses deterministic approximate text metrics. The same polygon intersection computes overlaps in both modes.

## Text projection

`word_document_project` renders a story as text. The same revision always
produces byte-identical output. Read the text, edit the text, send the changed
lines back through `word_document_patch`.

The projection hides the two offset spaces. Inspection offsets count rendered
characters and edit offsets count wire units. The anchor map is built in the
same pass as the text and holds both, so a patch never asks for a wire offset.

### Modes

| Mode | Content |
| --- | --- |
| `text` | Every paragraph of the story, one per line, including table cell paragraphs. No markup. |
| `md` | Headings, lists, GFM tables with cell content, and rich atom forms. Table cell paragraphs appear inside their table rows. |
| `outline` | Heading paragraphs only, as markdown headings. |

### Example

A brief with a heading, a mixed-formatting paragraph, a bullet list, a table,
an equation, and a closing section projects in `md` mode as:

```
# Findings
Decision: adopt the managed platform.
- Latency
- Cost

| Option | Score |
| --- | --- |
| Managed | 9 |

$E=mc^(2)$
## Next steps
Schedule the review.
```

The same document in `text` mode:

```
Findings
Decision: adopt the managed platform.
Latency
Cost
Option
Score
Managed
9
␏
Next steps
Schedule the review.
```

### Atom placeholders

Inline content with no editable text becomes one character in `text` mode, so
a projection column always names one whole atom.

| Content | `text` mode | `md` mode |
| --- | --- | --- |
| tab, positioned tab | `␉` U+2409 | `␉` |
| line break | `␊` U+240A | `␊` |
| page or column break | `␌` U+240C | `␌` |
| field | `␎` U+240E | `{{PAGE}}` from the instruction name |
| equation | `␏` U+240F | `$E=mc^(2)$` in linear math text |
| footnote or endnote reference | `␅` U+2405 | `[^3]` |
| image, chart, SmartArt, shape | `￼` U+FFFC | `![alt](object:12:3)` |
| ruby annotation | the base text | the base text |

The `object:*` target inside an `md` image form is a live reference. Pass it to
`word_document_inspect` with `"kind": "object"` for the full detail, or to a
drawing operation to edit it.

### Escaped lines

Plain paragraphs really do open with `3. ` or `- `. In `md` mode a paragraph
that carries no heading or list of its own, but whose text would read as a
marker, is written with a leading backslash:

```
\3. TERM AND TERMINATION
\- not a bullet
- Real bullet
```

The backslash is structure, not text. Keep it when you rewrite the line. Drop
it to turn the paragraph into a real list item. `text` mode carries no markdown
and no escapes.

Apart from that leading escape, `md` mode does not escape markdown characters
inside document text, and it does not emit inline emphasis. The projection is a
structural view for editing, not a markdown document to round-trip.

### Anchors

`anchors` holds one entry per projected line, in order.

```ts
interface AgentAnchorLine {
  line: number;                                   // 1-based within this window
  role: "paragraph" | "table" | "structure";
  blockRef?: string;
  marker: number;                                 // leading markdown structure characters
  editable: boolean;
  segments: AgentAnchorSegment[];
}

interface AgentAnchorSegment {
  start: number;                                  // columns in the line
  end: number;
  runRef: string;
  wireStart: number;                              // offsets in the run
  wireEnd: number;
  editable: boolean;                              // backed by one w:t
}
```

Only `paragraph` lines accept patches. A `table` line is the GFM rendering of a
whole table and a `structure` line is a blank separator; both carry context,
not edit targets. Patch table cells through `text` mode, which gives every cell
paragraph its own editable line.

The map is machine-facing. Read the text, not the anchors.

### Windows

A projection covers whole top-level blocks. When a window fills its budget the
result sets `truncated` and returns `next`; pass it back as `cursor` for the
following window. A cursor belongs to one revision and is rejected after any
edit. Every window carries the `revision` it was taken from.

## Patching a projection

A hunk replaces projection lines `startLine` through `endLine`, inclusive, with
the lines of `newText`. Line numbers are relative to the window, so repeat the
`story`, `mode`, and `cursor` that produced it. The refreshed projection in the
result covers that same window.

| Old lines | New lines | Result |
| --- | --- | --- |
| 1 | 1 | The differing span of the paragraph is rewritten. |
| 1 | many | The paragraph splits. Later lines inherit its style and list. |
| many | 1 | The paragraphs merge into the first, then the text is rewritten. |
| many | fewer | The extra paragraphs merge away. |

Delete a paragraph by covering it together with the line before it and
supplying only that line's text.

Rules:

- Every line a paragraph splits into repeats the first line's marker. Changing
  a marker on a split line is not supported.
- Changing the marker on a rewritten line changes the paragraph: `## ` sets the
  heading style, no marker clears it, and `- ` or `1. ` sets or clears the list.
  Changing a list indent is not supported.
- A hunk may not rewrite across an atom placeholder. Edit the atom with the
  operation that owns it.
- Hunks may not overlap.
- Text that replaces a span takes the formatting of the run where the span
  starts, the rule Word applies when you type over a mixed selection. Runs
  outside the changed span keep their own formatting.
- With `suggest: true` a hunk may add text or remove text, not both. Send a
  replacement as two patches.

Only the blocks a hunk touches must be unchanged. A patch written against an
older revision still applies when someone else edited a different part of the
document; it is rejected as stale once its own blocks move.

## References

- `block:*`: editable paragraph or table identity
- `run:*`: editable run identity
- `object:*`: inspectable component identity
- `asset:*`: readable or insertable binary asset identity
- `spatial:*`: one laid-out object occurrence
- `view:*`: revision-scoped inspection identity for content outside the current edit surface

Use references returned by the current document. Treat each reference as opaque. Re-inspect after structural edits.

## Editing

Every edit request contains the current revision and an array of operations. The result returns the new revision, the accepted operation kinds, the status, and the connection state.

Agent operations replace canonical internal IDs with these fields:

- `blockRef`: a paragraph target
- `runRef`: a run target
- `objectRef`: an inspectable drawing target returned as `editRef`
- `cellRef`: a paragraph inside a table cell
- `afterBlockRef`: a paragraph insertion target
- `at`: `{ blockRef, runRef, offset }`
- `assetRef`: an asset registered with `AgentDocument.addAsset` or supplied by the host

Internal IDs, client sequence values, node IDs, and provenance fields are allocated by the interface.

## Compose generic primitives

Use `word_document_compose` for new document structure. Use the combinations
below for later revisions and for loaded documents.

Use these combinations when one operation creates a target for later operations.

| Mechanism | Operation sequence |
| --- | --- |
| Add content to a new story | Call `ensureHeaderFooter`, inspect `overview`, read the returned story, then use its block and run references. |
| Repeat an object on section pages | Insert the object into a header or footer story. The story renders on each applicable page in that section. |
| Add an individual rule | Insert `line` or `verticalLine`. Use the inserted object's `editRef` with position, size, rotation, line-style, order, and removal operations. |
| Add a box around the page area | Use `setPageLayout.pageBorders`. This creates the section page-border structure. |
| Add a rule between text columns | Set `setPageLayout.columns` and `columnSeparator`. |
| Edit text inside a text box | Insert `textBox`, inspect the object, read its returned textbox story, then use text and paragraph operations on that story. |
| Edit an existing positioned object | Inspect its component and use `editRef`. Drawing operations apply through the object reference. |
| Edit Word-authored legacy rules | VML `v:line` objects inspect as `anchored-line`. Use the same drawing operations and `editRef` as a newly inserted line. |
| Change chart values | Insert a chart or inspect an existing chart, then use `setChartData` through its `editRef`. |
| Add binary media | Register bytes through `AgentDocument.addAsset` or a host asset facility, then pass the returned `assetRef` to `insertImage`. |

One header or footer XML story can produce several spatial occurrences. Spatial inspection returns one occurrence per rendered page with the same `editRef`.

## Complete operation catalog

The `word_document_capabilities` result is the authoritative closed schema. The tables below document every available operation and its agent-facing fields. Optional fields appear after `optional:`.

### Text

| Kind | Purpose | Fields |
| --- | --- | --- |
| `insertText` | Insert text at a position | `at`, `text`; optional: `suggest` |
| `deleteText` | Delete one run range | `blockRef`, `runRef`, `start`, `end` |
| `formatRun` | Format a complete run | `blockRef`, `runRef`, `patch` |
| `formatRange` | Format one run range | `blockRef`, `runRef`, `start`, `end`, `patch` |
| `setLink` | Add a safe hyperlink | `runRef`, `url` |
| `removeLink` | Remove a hyperlink | `runRef` |
| `toggleCheckbox` | Toggle a checkbox control | `runRef` |

### Paragraph

| Kind | Purpose | Fields |
| --- | --- | --- |
| `splitParagraph` | Split at a text position | `at`; optional: `suggest` |
| `mergeParagraph` | Merge into the previous paragraph | `blockRef` |
| `formatParagraph` | Set alignment or style | `blockRef`; optional: `align`, `styleId` |
| `setListType` | Set bullet, number, or normal paragraph | `blockRef`, `listKind` |
| `setListLevel` | Change list nesting | `blockRef`, `delta` |
| `adjustIndent` | Change indent by one step | `blockRef`, `direction` |
| `setSpacing` | Set line and paragraph spacing | `blockRef`, `patch` |
| `setDropCap` | Set or clear a drop cap | `blockRef`, `mode` |
| `setDivider` | Set or clear a paragraph divider | `blockRef`, `divider` |

### Review

| Kind | Purpose | Fields |
| --- | --- | --- |
| `commentRun` | Add a comment to a run | `runRef`, `text`; optional: `initials` |
| `replyComment` | Reply to a comment | `parentId`, `text`; optional: `initials` |
| `deleteComment` | Delete a comment thread | `commentId` |
| `suggestRevision` | Suggest text or paragraph deletion | optional: `ranges`, `marks` |
| `acceptRevision` | Accept one revision | `index` |
| `rejectRevision` | Reject one revision | `index` |
| `acceptAllRevisions` | Accept all revisions | none |
| `rejectAllRevisions` | Reject all revisions | none |

### Insert

| Kind | Purpose | Fields |
| --- | --- | --- |
| `pasteBlocks` | Insert validated Word paragraph blocks | `afterBlockRef`, `blocksXml` |
| `insertImage` | Insert a registered image | `runRef`, `assetRef`, `widthPx`, `heightPx` |
| `insertBreak` | Insert a page or column break | `runRef`, `breakKind` |
| `insertMath` | Insert native Word math | `runRef`, `mathText` |
| `insertShape` | Insert a shape or text box | `runRef`, `preset`; optional: `text` |
| `insertPageField` | Insert PAGE or PAGE/NUMPAGES | `runRef`, `fieldKind` |
| `insertFootnote` | Insert a footnote | `runRef`, `text` |
| `insertBookmark` | Insert a bookmark at a run | `runRef`, `name` |
| `insertBookmarkRange` | Bookmark a run range | `runRef`, `name`, `start`, `end` |
| `insertBlankPage` | Insert a blank page | `runRef` |
| `insertSectionBreak` | Insert a section boundary | `runRef`, `breakType` |
| `insertCrossRef` | Insert a bookmark reference | `runRef`, `bookmark`, `refKind` |
| `insertCoverPage` | Insert a native cover page | `content` |
| `insertWordArt` | Insert decorative WordArt | `runRef`, `text`, `preset` |
| `insertChart` | Insert a chart and workbook | `runRef`, `chart` |
| `insertSmartArt` | Insert a SmartArt diagram | `runRef`, `smartArt` |
| `insertDateTimeField` | Insert a DATE or TIME field | `runRef`, `dtKind`, `picture` |
| `insertField` | Insert an allowlisted Word field | `runRef`, `instruction`; optional: `cachedResult` |
| `insertTable` | Insert a table | `runRef`, `rows`, `cols` |

### Drawing

| Kind | Purpose | Fields |
| --- | --- | --- |
| `setDrawingRotation` | Rotate a drawing | `objectRef`, `degrees` |
| `setDrawingFill` | Set or clear fill | `objectRef`, `color` |
| `setDrawingLineStyle` | Set or clear shape outline or line style | `objectRef`, `color`; optional: `widthPx`, `dash` |
| `setDrawingOrder` | Move to front or back | `objectRef`, `order` |
| `setFloatingPagePosition` | Set page-relative position | `objectRef`, `xPx`, `yPx` |
| `resizeDrawing` | Set drawing dimensions | `objectRef`, `widthPx`, `heightPx` |
| `removeDrawing` | Remove a drawing | `objectRef` |
| `setImageAltText` | Set image accessibility text | `objectRef`, `alt` |
| `setImageWrap` | Set inline or floating wrapping | `objectRef`, `mode` |
| `setDrawingWordArtText` | Replace WordArt text | `objectRef`, `text` |
| `setDrawingWordArtStyle` | Set WordArt glyph color and opacity | `objectRef`, `color`, `opacity` |
| `setChartData` | Replace chart data | `objectRef`, `chart` |
| `setSmartArtData` | Replace SmartArt content | `objectRef`, `smartArt` |
| `setSmartArtNodeText` | Replace one SmartArt label | `objectRef`, `index`, `text` |
| `setSmartArtFill` | Set node or diagram fill | `objectRef`, `color`; optional: `nodeIndex` |
| `setSmartArtTextFormat` | Format SmartArt labels | `objectRef`, `format`; optional: `nodeIndex` |

Set `setDrawingLineStyle.color` to `null` to clear an outline. Supply
`widthPx` and `dash` when `color` contains an RGB value.

### Table

| Kind | Purpose | Fields |
| --- | --- | --- |
| `tableOp` | Add, delete, or format table parts | `cellRef`, `op` |
| `resizeTableColumn` | Move a column boundary | `cellRef`, `boundary`, `deltaPx`; optional: `renderedWidths` |
| `resizeTableRow` | Set row height | `cellRef`, `rowIdx`, `heightPx` |
| `moveTable` | Float and position a table | `cellRef`, `xPx`, `yPx`, `preservePageStart`, `pageDelta` |

### Math

| Kind | Purpose | Fields |
| --- | --- | --- |
| `setMathLinear` | Replace an equation | `blockRef`, `mathText` |
| `deleteMath` | Delete an equation | `blockRef` |
| `moveMath` | Move an equation | `blockRef`, `at` |

### Document

| Kind | Purpose | Fields |
| --- | --- | --- |
| `setPageLayout` | Set margins, page size, columns, or borders | `patch` |
| `setLineNumbering` | Configure margin line numbers | `patch` |
| `ensureHeaderFooter` | Create a header or footer story | `hfKind` |

## Nested value shapes

### Position

```json
{ "blockRef": "block:1", "runRef": "run:2", "offset": 0 }
```

### Run format patch

```ts
{
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: "#RRGGBB" | null;
  highlight?: string | null;
  fontSizePt?: number;
  fontFamily?: string;
  verticalAlign?: "superscript" | "subscript" | null;
  clear?: boolean;
}
```

### Paragraph spacing patch

```ts
{
  lineMultiple?: number;
  exactLinePt?: number;
  beforePt?: number | null;
  afterPt?: number | null;
}
```

### Paragraph divider

```ts
{
  style: "single" | "double" | "dotted" | "dashed" | "thinThickSmallGap";
  color: "#RRGGBB";
  widthPt: number;
  spacePt: number;
}
```

### Page layout patch

```ts
{
  margins?: { top?: number; right?: number; bottom?: number; left?: number }; // inches
  mirrorMargins?: boolean;
  size?: { width: number; height: number }; // inches
  orientation?: "portrait" | "landscape";
  columns?: number; // 1–12
  columnSeparator?: boolean;
  pageBorders?: { sz?: number; color?: string; offsetFrom?: "text" | "page" } | null;
}
```

### Cover page content

```ts
{ title: string; subtitle?: string; author?: string }
```

### Chart data

```ts
{
  type: "column" | "bar" | "line" | "pie";
  title?: string;
  categories: string[];
  series: Array<{ name: string; values: number[] }>;
}
```

Each series must contain the same number of values as `categories`.

### SmartArt data

```ts
{ layout: "process" | "cycle" | "hierarchy" | "list"; items: string[] }
```

### SmartArt text format

```ts
{
  fontFamily: string;
  fontSizePt: number;
  color: "RRGGBB";
  bold: boolean;
  italic: boolean;
  alignment: "left" | "center" | "right";
}
```

### Line numbering patch

```ts
{
  enabled: boolean;
  countBy?: number;
  restart?: "continuous" | "newPage" | "newSection";
  start?: number;
}
```

### Table operation

```ts
type TableOperation =
  | "deleteRow" | "deleteCol" | "deleteTable"
  | "rowAbove" | "rowBelow" | "colLeft" | "colRight"
  | { kind: "cellShading"; fill: "RRGGBB" | null }
  | { kind: "cellVAlign"; v: "top" | "center" | "bottom" }
  | { kind: "textWrapping"; wrapping: "none" | "around"; xPx: number; yPx: number };
```

### Suggested deletion

```ts
{
  ranges?: Array<{ blockRef: string; runRef: string; start: number; end: number }>;
  marks?: Array<{ blockRef: string; glyph: "ins" | "del" }>;
}
```

Set `suggest: true` on `insertText` or `splitParagraph` to create a tracked insertion. The interface adds the configured author and date.

### Fields

`insertField` accepts these field types as the first instruction token:

`PAGE`, `NUMPAGES`, `SECTIONPAGES`, `SECTION`, `DATE`, `TIME`, `CREATEDATE`, `SAVEDATE`, `PRINTDATE`, `AUTHOR`, `TITLE`, `SUBJECT`, `KEYWORDS`, `COMMENTS`, `FILENAME`, `NUMWORDS`, `NUMCHARS`, `PAGEREF`, `REF`, `SEQ`, `STYLEREF`, `TOC`, `INDEX`, `LISTNUM`, and `QUOTE`.

## JavaScript API

```ts
import { AgentDocument } from "@wordinweb/agent";

const document = AgentDocument.create({ provenance: { author: "Document agent" } });
const loaded = AgentDocument.load(docxBytes, { provenance: { author: "Document agent" } });
const connected = AgentDocument.connect(collaborativeTarget, { provenance: { author: "Document agent" } });

document.inspect(request);
document.capabilities(category?, kind?);
document.addAsset(bytes, mediaType);
document.asset(ref);
await document.compose({ revision: document.revision, body, header, footer, firstHeader, firstFooter, page });
await document.edit({ revision: document.revision, operations });
document.save();
document.tools();
```

`addAsset` returns an `asset:*` reference for `insertImage`. `tools()` returns the portable tool definitions described above.

## Session modes

- `AgentDocument.create` or `load`: headless local document
- `LocalDocumentSession`: one canonical document shared by a solo browser editor and an agent
- `collaborativeAgentTarget`: adapter for live, reconnecting, and durable offline collaboration sessions

The edit and inspection contracts stay the same in every mode.
