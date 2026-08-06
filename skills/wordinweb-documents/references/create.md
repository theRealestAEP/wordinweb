# Create a DOCX

Use this reference for a new document. It contains the complete common creation
contract. Open [interface.md](interface.md) only for a type that this reference
does not include.

Start a JavaScript authoring program with this import:

```js
import { AgentDocument } from "@wordinweb/agent";
```

## Workflow

1. Create an `AgentDocument` or use `word_document_compose` from the host.
2. Register each image and keep its `asset:*` reference.
3. Compose the body and repeating page furniture in one call.
4. Check `result.overview`, `result.createdObjects`, and `result.createdObjectsTruncated`.
5. Inspect the required page range with `spatial`.
6. Inspect an object only when its compact summary lacks a required fact.
7. Save the DOCX.
8. Render and inspect every page when the host supplies a DOCX renderer.

Use `doc.addAsset(bytes, mediaType)` for each image. Pass a `Uint8Array` or
`Buffer` as `bytes`. The method returns the `asset:*` reference for an image
block.

In standalone Node.js, save with `await writeFile("output.docx", doc.save())`.
`doc.save()` returns the DOCX as a `Uint8Array`. Treat rendering as a host
facility. Use semantic and spatial inspection when the standalone workspace
does not supply a renderer.

## Compose request

```ts
interface AgentComposeRequest {
  revision: string;
  page?: PageLayoutPatch;
  body: AgentComposeBlock[];
  header?: AgentComposeBlock[];
  footer?: AgentComposeBlock[];
  firstHeader?: AgentComposeBlock[];
  firstFooter?: AgentComposeBlock[];
}
```

`firstHeader` and `firstFooter` enable Word's different-first-page layout.
Use them when a cover background needs different contrast from later pages.

Page sizes and page margins use inches. Letter size is `{ width: 8.5, height:
11 }`. Drawing sizes and drawing positions use pixels. Text sizes use points.

```ts
type AgentComposeBlock =
  | { type: "paragraph"; text?: string; runs?: AgentComposeRun[]; styleId?: "Normal" | "Title" | "Subtitle" | "Heading1" | "Heading2" | "Heading3"; align?: "left" | "center" | "right" | "both"; spacing?: { beforePt?: number; afterPt?: number; lineMultiple?: number }; list?: "bullet" | "number" }
  | { type: "heading"; level: 1 | 2 | 3; text: string }
  | { type: "table"; rows: Array<Array<string | AgentComposeCell>>; headerRows?: number; headerFill?: string; headerTextColor?: string; columnWidths?: number[] }
  | { type: "equation"; mathText: string; align?: "left" | "center" | "right" }
  | { type: "chart"; chart: ChartData; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "smartArt"; smartArt: SmartArtData; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "image"; assetRef: string; widthPx: number; heightPx: number; alt?: string; align?: "left" | "center" | "right"; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; position?: { xPx: number; yPx: number } }
  | { type: "shape"; preset: "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox"; text?: string; textStyle?: AgentComposeTextStyle; widthPx?: number; heightPx?: number; position?: { xPx: number; yPx: number }; fill?: string | null; line?: DrawingLine | null; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; order?: "front" | "back" }
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

interface DrawingLine {
  color: string;
  widthPx: number;
  dash: "solid" | "dashed" | "dotted";
}

interface ChartData {
  type: "column" | "bar" | "line" | "pie";
  title?: string;
  categories: string[];
  series: Array<{ name: string; values: number[] }>;
}

interface SmartArtData {
  layout: "process" | "cycle" | "hierarchy" | "list";
  items: string[];
}
```

Use either `text` or `runs` in a paragraph or table cell. Use `textStyle` for
shape text. Set `line: null` for a borderless shape. Supply both drawing
dimensions when you supply either one. Use `heading` for the native outline.
Consecutive list paragraphs share one Word numbering sequence.

An unpositioned shape uses `topAndBottom` wrap. A positioned shape uses `none`.
Set `wrap` to override either default.

Floating drawing positions use page coordinates. For a full Letter-page
background, use `widthPx: 816`, `heightPx: 1056`, and `position: { xPx: 0,
yPx: 0 }`.

For a watermark, prefer the `insertWatermark` edit operation on an existing
document. It writes Word's own VML stamp into every header part, so Word and
this renderer draw the same thing. The recipe below is for a decorative
background that is not a watermark: place WordArt in `header` with
`wrap: "behind"`, `order: "back"`, and an opacity near `0.05` to `0.12`.

## Compact compose result

```ts
interface AgentComposeResult {
  revision: string;
  status: "applied";
  blocks: number;
  components: number;
  overview: AgentOverview;
  createdObjects: Array<{
    ref: string;
    editRef?: string;
    type: string;
    label?: string;
  }>;
  createdObjectsTruncated: boolean;
}
```

`overview.objectCounts` reports semantic object types. `overview.components`
reports raw run-content kinds. Use `createdObjects` to locate fields, images,
equations, shapes, WordArt, charts, and SmartArt without a full story read.

The semantic type describes the saved DOCX object:

- An inline image has type `image`.
- A positioned image has type `anchored-image`.
- A shape or WordArt object has type `anchored-textbox`. Its label contains its
  text when text exists.
- A page number has type `field`. `pageOfTotal` creates `PAGE` and `NUMPAGES`
  field objects.
- A page break has type `break`.

Use the label and `editRef` to distinguish objects that share a semantic type.

## Example

```ts
const doc = AgentDocument.create();
const result = await doc.compose({
  revision: doc.revision,
  page: {
    size: { width: 8.5, height: 11 },
    margins: { top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 },
  },
  body: [
    { type: "paragraph", text: "Launch guide", styleId: "Title" },
    { type: "heading", level: 1, text: "Opening checklist" },
    { type: "paragraph", text: "Confirm the room.", list: "number" },
    { type: "paragraph", text: "Brief the team.", list: "number" },
    { type: "shape", preset: "rectangle", text: "Team note", textStyle: { color: "FFFFFF", bold: true, align: "center" }, widthPx: 240, heightPx: 80, fill: "10243E", line: null },
  ],
  header: [{ type: "wordArt", text: "DRAFT", preset: "plain", opacity: 0.08, wrap: "behind", order: "back" }],
  footer: [{ type: "pageNumber", fieldKind: "pageOfTotal", align: "center", color: "10243E" }],
  firstFooter: [{ type: "pageNumber", fieldKind: "pageOfTotal", align: "center", color: "FFFFFF" }],
});

doc.inspect({ kind: "spatial", pages: { start: 1, count: 10 }, includeOverlaps: true });
```
