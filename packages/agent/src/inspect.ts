import {
  imageAltText,
  runWireLength,
  type Block,
  type DocxDocument,
  type MathNode,
  type Paragraph,
  type Run,
  type RunContent,
  type Shape,
  type StableIds,
} from "@wordinweb/core";
import { assetRef, blockObjectRef, blockRef, objectRef, parseObjectRef, runRef, viewRef } from "./refs.js";
import type {
  AgentAsset,
  AgentBlock,
  AgentComponentSummary,
  AgentContextParagraph,
  AgentContextResult,
  AgentObjectResult,
  AgentOverview,
  AgentParagraph,
  AgentReadResult,
  AgentRun,
  AgentSearchResult,
} from "./types.js";

interface Story {
  id: string;
  kind: "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";
  blocks: Block[];
}

interface IndexedBlock {
  block: Block;
  story: string;
  ref: string;
  editable: boolean;
}

interface ObjectEntry {
  editRef?: string;
  runId?: number;
  objectIndex?: number;
  type: string;
  detail: Record<string, unknown>;
}

export function mathText(nodes: MathNode[]): string {
  const render = (node: MathNode): string => {
    switch (node.t) {
      case "run": return node.text;
      case "sup": return `${mathText(node.base)}^(${mathText(node.script)})`;
      case "sub": return `${mathText(node.base)}_(${mathText(node.script)})`;
      case "frac": return `(${mathText(node.num)})/(${mathText(node.den)})`;
      case "rad": return `root(${node.deg ? `${mathText(node.deg)},` : ""}${mathText(node.e)})`;
      case "nary": return `${node.chr}[${mathText(node.sub)},${mathText(node.sup)}](${mathText(node.e)})`;
      case "dlm": return `${node.beg}${node.e.map(mathText).join("|")}${node.end}`;
      case "mat": return `[${node.rows.map((row) => row.map(mathText).join(",")).join(";")}]`;
      case "eqarr": return node.rows.map(mathText).join(";");
      case "acc": return `${node.chr}(${mathText(node.e)})`;
      case "grp": return `${node.chr}(${mathText(node.e)})`;
      case "lim": return `${mathText(node.e)}_${node.pos}(${mathText(node.lim)})`;
    }
  };
  return nodes.map(render).join("");
}

function plainMath(nodes: MathNode[]): MathNode[] {
  return JSON.parse(JSON.stringify(nodes)) as MathNode[];
}

function mediaType(part: string): string {
  const ext = part.split(".").pop()?.toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "bmp") return "image/bmp";
  if (ext === "svg") return "image/svg+xml";
  return "application/octet-stream";
}

function blockText(block: Block): string {
  if (block.type === "paragraph") return paragraphText(block);
  return block.rows.flatMap((row) => row.cells).flatMap((cell) => cell.blocks).map(blockText).join("\n");
}

export function runText(run: Run): string {
  let out = "";
  for (const content of run.content) {
    switch (content.kind) {
      case "text": out += content.text; break;
      case "break": out += "\n"; break;
      case "tab":
      case "ptab": out += "\t"; break;
      case "field": out += content.cachedResult; break;
      case "math": out += mathText(content.nodes); break;
      case "ruby": out += runText(content.base); break;
      case "image":
      case "anchor":
      case "drawing":
      case "noteRef": break;
      default: {
        const exhaustive: never = content;
        return exhaustive;
      }
    }
  }
  return out;
}

export function paragraphRuns(paragraph: Paragraph): Array<{ run: Run; hyperlink?: { href?: string; anchor?: string } }> {
  const runs: Array<{ run: Run; hyperlink?: { href?: string; anchor?: string } }> = [];
  for (const child of paragraph.children) {
    if (child.type === "run") runs.push({ run: child });
    else for (const run of child.runs) runs.push({ run, hyperlink: { href: child.href, anchor: child.anchor } });
  }
  return runs;
}

function paragraphText(paragraph: Paragraph): string {
  return paragraphRuns(paragraph).map(({ run }) => runText(run)).join("");
}

function shapeBase(shape: Shape): Record<string, unknown> {
  const base: Record<string, unknown> = {
    shapeType: shape.type,
    zOrder: shape.z,
  };
  if (shape.type === "line") {
    return { ...base, x1: shape.x1, y1: shape.y1, x2: shape.x2, y2: shape.y2, color: shape.color, weight: shape.weight, style: shape.style ?? "single", hRel: shape.hRel, vRel: shape.vRel };
  }
  return {
    ...base,
    x: shape.x,
    y: shape.y,
    width: shape.width,
    height: shape.height,
    hRel: shape.hRel,
    vRel: shape.vRel,
    hAlign: shape.hAlign,
    vAlign: shape.vAlign,
    rotation: shape.rotation,
    behind: shape.behind,
  };
}

export class SemanticInspector {
  private readonly stories = new Map<string, Story>();
  private readonly entries = new Map<string, IndexedBlock[]>();
  private readonly objects = new Map<string, ObjectEntry>();
  private readonly assetsByRef = new Map<string, string>();
  private readonly assetsByPart = new Map<string, string>();
  private readonly components: Record<string, number> = {};
  private readonly objectCounts: Record<string, number> = {};
  private readonly indexedBlocks = new WeakMap<Block, IndexedBlock>();
  private readonly componentsByRun = new WeakMap<Run, AgentComponentSummary[]>();
  private readonly refsByRun = new WeakMap<Run, { ref: string; editable: boolean }>();
  private viewBlockIndex = 0;
  private viewRunIndex = 0;
  private viewObjectIndex = 0;

  constructor(
    private readonly doc: DocxDocument,
    private readonly ids: StableIds,
    private readonly revision: string,
  ) {
    this.addStory({ id: "body", kind: "body", blocks: doc.sections.flatMap((section) => section.blocks) });
    for (const [id, value] of doc.headers) this.addStory({ id: `header:${id}`, kind: "header", blocks: value.blocks });
    for (const [id, value] of doc.footers) this.addStory({ id: `footer:${id}`, kind: "footer", blocks: value.blocks });
    for (const [id, blocks] of doc.footnotes) this.addStory({ id: `footnote:${id}`, kind: "footnote", blocks });
    for (const [id, blocks] of doc.endnotes) this.addStory({ id: `endnote:${id}`, kind: "endnote", blocks });
    this.indexStories();
  }

  private addStory(story: Story): void {
    if (!this.stories.has(story.id)) this.stories.set(story.id, story);
  }

  private indexStories(): void {
    const queue = [...this.stories.values()];
    for (let storyIndex = 0; storyIndex < queue.length; storyIndex++) {
      const story = queue[storyIndex];
      const indexed: IndexedBlock[] = [];
      const visit = (blocks: Block[]): void => {
        for (const block of blocks) {
          const stableId = block.src ? this.ids.idOf(block.src) : undefined;
          const entry: IndexedBlock = {
            block,
            story: story.id,
            ref: stableId === undefined ? viewRef(this.revision, "block", ++this.viewBlockIndex) : blockRef(stableId),
            editable: stableId !== undefined,
          };
          indexed.push(entry);
          this.indexedBlocks.set(block, entry);
          if (block.type === "table") {
            for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
          } else {
            this.indexParagraphObjects(block, entry, queue);
          }
        }
      };
      visit(story.blocks);
      this.entries.set(story.id, indexed);
    }
  }

  private count(kind: string): void {
    this.components[kind] = (this.components[kind] ?? 0) + 1;
  }

  private registerAsset(part: string): string {
    const current = this.assetsByPart.get(part);
    if (current) return current;
    const ref = assetRef(this.assetsByRef.size + 1);
    this.assetsByPart.set(part, ref);
    this.assetsByRef.set(ref, part);
    return ref;
  }

  private addTextboxStory(ref: string, suffix: string, blocks: Block[], queue: Story[]): string {
    const id = `textbox:${ref}:${suffix}`;
    if (!this.stories.has(id)) {
      const story: Story = { id, kind: "textbox", blocks };
      this.stories.set(id, story);
      queue.push(story);
    }
    return id;
  }

  private indexParagraphObjects(paragraph: Paragraph, block: IndexedBlock, queue: Story[]): void {
    paragraphRuns(paragraph).forEach(({ run }, runIndex) => {
      const runId = run.src ? this.ids.idOf(run.src) : undefined;
      this.refsByRun.set(run, {
        ref: runId === undefined ? viewRef(this.revision, "run", ++this.viewRunIndex) : runRef(runId),
        editable: runId !== undefined,
      });
      const summaries: AgentComponentSummary[] = [];
      run.content.forEach((content, index) => {
        this.count(content.kind);
        if (content.kind === "text") return;
        const stableBlockId = block.editable ? Number(block.ref.slice("block:".length)) : undefined;
        const ref = runId !== undefined
          ? objectRef(runId, index)
          : stableBlockId !== undefined
            ? blockObjectRef(stableBlockId, runIndex, index)
            : viewRef(this.revision, "object", ++this.viewObjectIndex);
        const editRef = content.kind === "math" && block.editable
          ? block.ref
          : runId !== undefined && (content.kind === "image" || content.kind === "anchor" || content.kind === "drawing")
            ? ref
            : runId !== undefined
              ? runRef(runId)
              : undefined;
        const entry = this.objectEntry(content, ref, editRef, queue);
        entry.runId = runId;
        entry.objectIndex = index;
        this.objects.set(ref, entry);
        this.objectCounts[entry.type] = (this.objectCounts[entry.type] ?? 0) + 1;
        summaries.push(this.objectSummary(ref, entry));
      });
      this.componentsByRun.set(run, summaries);
    });
  }

  private objectSummary(ref: string, entry: ObjectEntry): AgentComponentSummary {
    const detail = entry.detail;
    const chart = detail.chart && typeof detail.chart === "object" ? detail.chart as { title?: unknown } : undefined;
    const smartArt = detail.smartArt && typeof detail.smartArt === "object" ? detail.smartArt as { layout?: unknown } : undefined;
    const label = [detail.linear, detail.text, detail.altText, detail.instruction, chart?.title, smartArt?.layout, detail.breakType]
      .find((value): value is string => typeof value === "string" && value.length > 0);
    return { ref, editRef: entry.editRef, type: entry.type, ...(label ? { label: label.slice(0, 200) } : {}) };
  }

  private objectEntry(content: Exclude<RunContent, { kind: "text" }>, ref: string, editRef: string | undefined, queue: Story[]): ObjectEntry {
    switch (content.kind) {
      case "break": return { editRef, type: "break", detail: { breakType: content.breakType } };
      case "tab": return { editRef, type: "tab", detail: {} };
      case "ptab": return { editRef, type: "positioned-tab", detail: { alignment: content.alignment, relativeTo: content.relativeTo } };
      case "field": return { editRef, type: "field", detail: { instruction: content.instruction, cachedResult: content.cachedResult, checkbox: Boolean(content.checkbox) } };
      case "math": return { editRef, type: "math", detail: { linear: mathText(content.nodes), nodes: plainMath(content.nodes), display: Boolean(content.display), justification: content.jc } };
      case "noteRef": return { editRef, type: `${content.noteType}-reference`, detail: { id: content.id, self: Boolean(content.self), customMarkFollows: Boolean(content.customMarkFollows) } };
      case "ruby": return { editRef, type: "ruby", detail: { base: runText(content.base), annotation: runText(content.rt), raiseHalfPoints: content.hpsRaise, alignment: content.align } };
      case "image": {
        const detail: Record<string, unknown> = {
          assetRef: this.registerAsset(content.part), width: content.width, height: content.height,
          anchored: Boolean(content.anchored), crop: content.crop, rotation: content.rotation,
          altText: content.srcDrawing ? imageAltText(content.srcDrawing) : "",
        };
        if (content.model3D) detail.model3D = { assetRef: this.registerAsset(content.model3D.part), posterAssetRef: this.registerAsset(content.model3D.posterPart), rotation: content.model3D.rotation };
        if (content.webVideo) detail.webVideo = { url: content.webVideo.url, width: content.webVideo.width, height: content.webVideo.height };
        if (content.embeddedObject) detail.embeddedObject = { assetRef: this.registerAsset(content.embeddedObject.part), filename: content.embeddedObject.filename, progId: content.embeddedObject.progId };
        return { editRef, type: "image", detail };
      }
      case "anchor": {
        const shape = content.shape;
        const detail = shapeBase(shape);
        if (shape.type === "image") {
          detail.assetRef = this.registerAsset(shape.part);
          detail.wrap = shape.wrap;
          detail.crop = shape.crop;
          detail.altText = shape.srcDrawing ? imageAltText(shape.srcDrawing) : "";
        } else if (shape.type === "textbox") {
          detail.story = this.addTextboxStory(ref, "text", shape.blocks, queue);
          detail.text = shape.blocks.map(blockText).join("\n").trim();
          detail.wrap = shape.wrap;
          detail.fill = shape.fill;
          detail.stroke = shape.stroke;
          detail.allowOverlap = shape.allowOverlap !== false;
          detail.wordArt = Boolean(shape.wordArt);
          if (shape.wordArt) {
            detail.fill = shape.wordArtFill ?? detail.fill;
            detail.opacity = shape.wordArtOpacity ?? 1;
          }
          detail.warp = shape.warp;
        } else if (shape.type === "art") {
          detail.lines = shape.lines;
          detail.paths = shape.paths;
          detail.images = shape.images.map((image) => ({ ...image, assetRef: this.registerAsset(image.part), part: undefined }));
          detail.textStories = (shape.texts ?? []).map((text, index) => this.addTextboxStory(ref, `text-${index}`, text.blocks, queue));
          detail.ink = Boolean(shape.ink);
        } else if (shape.type === "wordart") {
          Object.assign(detail, { wordArt: true, text: shape.text, fontFamily: shape.fontFamily, bold: shape.bold, italic: shape.italic, fill: shape.fill, opacity: shape.opacity, noFit: shape.noFit });
        }
        return { editRef, type: `anchored-${shape.type}`, detail };
      }
      case "drawing": {
        const detail: Record<string, unknown> = {
          width: content.width,
          height: content.height,
          lines: content.lines,
          paths: content.paths ?? [],
          images: content.images.map((image) => ({ ...image, assetRef: this.registerAsset(image.part), part: undefined })),
          chart: content.chart,
          smartArt: content.smartArt,
        };
        if (content.textbox) detail.textboxStory = this.addTextboxStory(ref, "text", content.textbox.blocks, queue);
        detail.textStories = (content.texts ?? []).map((text, index) => this.addTextboxStory(ref, `text-${index}`, text.blocks, queue));
        const type = content.chart ? "chart" : content.smartArt ? "smart-art" : content.textbox ? "text-box" : "drawing";
        return { editRef, type, detail };
      }
      default: {
        const exhaustive: never = content;
        return exhaustive;
      }
    }
  }

  overview(): AgentOverview {
    let paragraphs = 0;
    let tables = 0;
    let characters = 0;
    const outline: AgentOverview["outline"] = [];
    for (const entry of this.entries.get("body") ?? []) {
      if (entry.block.type === "table") {
        tables++;
        continue;
      }
      paragraphs++;
      const text = paragraphText(entry.block);
      characters += text.length;
      const effective = this.doc.effectiveParaProps(entry.block);
      const level = effective.outlineLevel;
      if (level !== undefined && entry.block.src) {
        const id = this.ids.idOf(entry.block.src);
        if (id !== undefined) outline.push({ ref: blockRef(id), level: level + 1, text: text.trim(), styleId: entry.block.props.styleId });
      }
    }
    return {
      revision: this.revision,
      sections: this.doc.sections.length,
      mirrorMargins: this.doc.mirrorMargins,
      sectionLayouts: this.doc.sections.map((section, index) => ({
        index,
        page: { width: section.props.pageWidth, height: section.props.pageHeight },
        margins: {
          top: section.props.marginTop,
          right: section.props.marginRight,
          bottom: section.props.marginBottom,
          left: section.props.marginLeft,
        },
        headerDistance: section.props.headerDistance,
        footerDistance: section.props.footerDistance,
        gutter: section.props.gutter,
        columns: {
          count: section.props.columns.count,
          space: section.props.columns.space,
          widths: section.props.columns.widths,
          spaces: section.props.columns.spaces,
          separator: section.props.columns.sep,
        },
        titlePage: section.props.titlePage,
        pageNumberStart: section.props.pageNumberStart,
        pageNumberFormat: section.props.pageNumberFormat,
        breakType: section.props.type,
        verticalAlignment: section.props.vAlign,
        pageBorders: section.props.pageBorders ? JSON.parse(JSON.stringify(section.props.pageBorders)) : undefined,
        // section.props.lineNumbering.start is the raw w:start OFFSET (see
        // core/model.ts); report the human-meaningful first-line number
        // instead, matching what setLineNumbering's own patch.start expects.
        lineNumbering: section.props.lineNumbering
          ? { ...section.props.lineNumbering, start: section.props.lineNumbering.start + 1 }
          : undefined,
      })),
      stories: [...this.stories.values()].map((story) => ({ id: story.id, kind: story.kind, blocks: this.entries.get(story.id)?.length ?? 0 })),
      blocks: { paragraphs, tables },
      characters,
      comments: this.doc.comments.length,
      components: { ...this.components },
      objectCounts: { ...this.objectCounts },
      outline,
    };
  }

  context(
    storyIds?: string[],
    maxBlocks = 100,
    maxCharacters = 24_000,
    include: Array<"bookmarks" | "objects"> = [],
    includeEmpty = false,
  ): AgentContextResult {
    if (storyIds !== undefined && (!Array.isArray(storyIds) || storyIds.length < 1 || storyIds.length > 100)) {
      throw new Error("stories must contain 1 to 100 story IDs");
    }
    if (!Number.isInteger(maxBlocks) || maxBlocks < 1 || maxBlocks > 200) throw new Error("maxBlocks must be between 1 and 200");
    if (!Number.isInteger(maxCharacters) || maxCharacters < 1 || maxCharacters > 100_000) throw new Error("maxCharacters must be between 1 and 100000");
    if (!Array.isArray(include) || include.some((value) => value !== "bookmarks" && value !== "objects")) {
      throw new Error("include supports bookmarks and objects");
    }
    const selected = storyIds ?? [...this.stories.keys()];
    if (new Set(selected).size !== selected.length) throw new Error("stories must not contain duplicates");
    for (const id of selected) if (!this.stories.has(id)) throw new Error(`Unknown story: ${id}`);

    const available = selected.filter((id) => (this.entries.get(id) ?? []).some((entry) =>
      entry.block.type === "table" || includeEmpty || paragraphText(entry.block).length > 0));
    const contents: AgentContextResult["contents"] = [];
    let usedBlocks = 0;
    let usedCharacters = 0;
    let remainingStories: string[] | undefined;

    for (let storyIndex = 0; storyIndex < available.length; storyIndex++) {
      const storyId = available[storyIndex];
      const story = this.stories.get(storyId)!;
      const entries = this.entries.get(storyId) ?? [];
      const blocks: AgentContextResult["contents"][number]["blocks"] = [];
      let next: AgentReadResult["next"];

      for (let index = 0; index < entries.length; index++) {
        const entry = entries[index];
        if (entry.block.type === "paragraph" && !includeEmpty && paragraphText(entry.block).length === 0) continue;
        if (usedBlocks >= maxBlocks || (entry.block.type === "paragraph" && usedCharacters >= maxCharacters)) {
          next = { value: `cursor:${this.revision}:${index}:0` };
          break;
        }
        if (entry.block.type === "table") {
          blocks.push({
            type: "table",
            ref: entry.ref,
            rows: entry.block.rows.length,
            columns: Math.max(0, ...entry.block.rows.map((row) => row.cells.length)),
          });
          usedBlocks++;
          continue;
        }
        const text = paragraphText(entry.block);
        const end = Math.min(text.length, maxCharacters - usedCharacters);
        blocks.push(this.contextParagraph(entry, 0, end, include));
        usedBlocks++;
        usedCharacters += end;
        if (end < text.length) {
          next = { value: `cursor:${this.revision}:${index}:${end}` };
          break;
        }
      }

      if (next && blocks.length === 0) {
        remainingStories = available.slice(storyIndex);
        break;
      }
      if (blocks.length > 0) contents.push({ story: storyId, kind: story.kind, blocks, ...(next ? { next } : {}) });
      if (next) {
        remainingStories = available.slice(storyIndex + 1);
        break;
      }
    }

    const truncated = contents.some((story) => story.next !== undefined) || Boolean(remainingStories?.length);
    return {
      revision: this.revision,
      contents,
      truncated,
      ...(remainingStories?.length ? { remainingStories } : {}),
    };
  }

  private contextParagraph(
    indexed: IndexedBlock,
    start: number,
    end: number,
    include: Array<"bookmarks" | "objects">,
  ): AgentContextParagraph {
    const paragraph = indexed.block as Paragraph;
    const allText = paragraphText(paragraph);
    const runs: AgentContextParagraph["runs"] = [];
    const objects: AgentComponentSummary[] = [];
    let paragraphOffset = 0;
    for (const { run } of paragraphRuns(paragraph)) {
      const text = runText(run);
      const runStart = paragraphOffset;
      const runEnd = runStart + text.length;
      paragraphOffset = runEnd;
      if (include.includes("objects")) objects.push(...(this.componentsByRun.get(run) ?? []));
      const overlaps = text.length === 0 ? start === 0 && end === 0 : runEnd > start && runStart < end;
      if (!overlaps) continue;
      const identity = this.refsByRun.get(run)!;
      const wireLength = run.src ? runWireLength(run.src) : text.length;
      runs.push({ ref: identity.ref, start: runStart, end: runEnd, ...(wireLength !== text.length ? { wireLength } : {}) });
    }
    const effective = this.doc.effectiveParaProps(paragraph);
    const bookmarks = paragraph.bookmarks ?? [];
    return {
      type: "paragraph",
      ref: indexed.ref,
      text: allText.slice(start, end),
      ...(start > 0 || end < allText.length ? { range: { start, end, total: allText.length } } : {}),
      runs,
      ...(paragraph.props.styleId ? { styleId: paragraph.props.styleId } : {}),
      ...(effective.outlineLevel === undefined ? {} : { outlineLevel: effective.outlineLevel + 1 }),
      ...(effective.numbering ? { list: { numId: effective.numbering.numId, level: effective.numbering.ilvl } } : {}),
      ...(include.includes("bookmarks") && bookmarks.length ? { bookmarks } : {}),
      ...(include.includes("objects") && objects.length ? { objects } : {}),
    };
  }

  read(storyId = "body", cursorValue?: string, maxBlocks = 20, maxCharacters = 12_000): AgentReadResult {
    if (!Number.isInteger(maxBlocks) || maxBlocks < 1 || maxBlocks > 200) throw new Error("maxBlocks must be between 1 and 200");
    if (!Number.isInteger(maxCharacters) || maxCharacters < 1 || maxCharacters > 100_000) throw new Error("maxCharacters must be between 1 and 100000");
    const entries = this.entries.get(storyId);
    if (!entries) throw new Error("Unknown story");
    let index = 0;
    let textOffset = 0;
    if (cursorValue) {
      const match = /^cursor:([^:]+):(\d+):(\d+)$/.exec(cursorValue);
      if (!match || match[1] !== this.revision) throw new Error("Invalid or stale cursor");
      index = Number(match[2]);
      textOffset = Number(match[3]);
    }
    const blocks: AgentBlock[] = [];
    let used = 0;
    while (index < entries.length && blocks.length < maxBlocks && used < maxCharacters) {
      const entry = entries[index];
      const current = entry.block;
      if (current.type === "table") {
        blocks.push(this.table(entry));
        index++;
        textOffset = 0;
        continue;
      }
      const total = paragraphText(current).length;
      const available = maxCharacters - used;
      const end = Math.min(total, textOffset + available);
      blocks.push(this.paragraph(entry, textOffset, end));
      used += end - textOffset;
      if (end < total) {
        textOffset = end;
        break;
      }
      index++;
      textOffset = 0;
    }
    const truncated = index < entries.length || textOffset > 0;
    return {
      revision: this.revision,
      story: storyId,
      blocks,
      ...(truncated ? { next: { value: `cursor:${this.revision}:${index}:${textOffset}` } } : {}),
      truncated,
    };
  }

  private paragraph(indexed: IndexedBlock, start: number, end: number): AgentParagraph {
    const paragraph = indexed.block as Paragraph;
    const effectiveParagraph = this.doc.effectiveParaProps(paragraph);
    const allText = paragraphText(paragraph);
    const runs: AgentRun[] = [];
    let paragraphOffset = 0;
    for (const { run, hyperlink } of paragraphRuns(paragraph)) {
      const text = runText(run);
      const runStart = paragraphOffset;
      const runEnd = runStart + text.length;
      paragraphOffset = runEnd;
      if (runEnd < start || runStart > end) continue;
      const runId = run.src ? this.ids.idOf(run.src) : undefined;
      const localStart = Math.max(0, start - runStart);
      const localEnd = Math.min(text.length, end - runStart);
      const components = this.componentsByRun.get(run) ?? [];
      const runIdentity = this.refsByRun.get(run)!;
      runs.push({
        ref: runIdentity.ref,
        editable: runIdentity.editable,
        paragraphStart: runStart,
        paragraphEnd: runEnd,
        wireLength: runId !== undefined && run.src ? runWireLength(run.src) : text.length,
        text: text.slice(localStart, Math.max(localStart, localEnd)),
        formatting: { ...this.doc.effectiveRunProps(paragraph, run.props) },
        ...(hyperlink ? { hyperlink } : {}),
        components,
      });
    }
    return {
      type: "paragraph",
      ref: indexed.ref,
      editable: indexed.editable,
      story: indexed.story,
      styleId: paragraph.props.styleId,
      outlineLevel: effectiveParagraph.outlineLevel === undefined ? undefined : effectiveParagraph.outlineLevel + 1,
      alignment: effectiveParagraph.alignment,
      list: effectiveParagraph.numbering ? { numId: effectiveParagraph.numbering.numId, level: effectiveParagraph.numbering.ilvl } : effectiveParagraph.numbering === null ? null : undefined,
      text: allText.slice(start, end),
      textRange: { start, end, total: allText.length },
      runs,
      bookmarks: paragraph.bookmarks ?? [],
    };
  }

  private table(indexed: IndexedBlock): AgentBlock {
    const table = indexed.block as Extract<Block, { type: "table" }>;
    return {
      type: "table",
      ref: indexed.ref,
      editable: indexed.editable,
      story: indexed.story,
      rows: table.rows.length,
      columns: Math.max(0, ...table.rows.map((row) => row.cells.length)),
      styleId: table.props.styleId,
      floating: Boolean(table.props.floating),
      cells: table.rows.flatMap((row, rowIndex) => row.cells.map((cell, columnIndex) => ({
        row: rowIndex,
        column: columnIndex,
        blocks: cell.blocks.flatMap((block) => {
          const entry = this.indexedBlocks.get(block);
          return entry ? [entry.ref] : [];
        }),
      }))),
    };
  }

  search(query: string, maxResults = 50): AgentSearchResult {
    if (typeof query !== "string" || query.length === 0 || query.length > 1_000) throw new Error("query must contain 1 to 1000 characters");
    if (!Number.isInteger(maxResults) || maxResults < 1 || maxResults > 500) throw new Error("maxResults must be between 1 and 500");
    const needle = query.toLocaleLowerCase();
    const matches: AgentSearchResult["matches"] = [];
    for (const entries of this.entries.values()) {
      for (const indexed of entries) {
        const block = indexed.block;
        if (block.type !== "paragraph") continue;
        for (const { run } of paragraphRuns(block)) {
          const runId = run.src ? this.ids.idOf(run.src) : undefined;
          const runIdentity = this.refsByRun.get(run)!;
          const text = runText(run);
          const lower = text.toLocaleLowerCase();
          let at = lower.indexOf(needle);
          while (at >= 0) {
            matches.push({ blockRef: indexed.ref, runRef: runIdentity.ref, editable: indexed.editable && runId !== undefined, start: at, end: at + query.length, excerpt: text.slice(Math.max(0, at - 60), Math.min(text.length, at + query.length + 60)) });
            if (matches.length >= maxResults) return { revision: this.revision, matches, truncated: true };
            at = lower.indexOf(needle, at + Math.max(1, needle.length));
          }
        }
      }
    }
    return { revision: this.revision, matches, truncated: false };
  }

  object(ref: string): AgentObjectResult {
    parseObjectRef(ref);
    const entry = this.objects.get(ref);
    if (!entry) throw new Error("Unknown object reference");
    return { revision: this.revision, ref, editRef: entry.editRef, type: entry.type, detail: entry.detail };
  }

  objectCatalog(maxObjects = 100): { items: AgentComponentSummary[]; truncated: boolean } {
    const all = [...this.objects].map(([ref, entry]) => this.objectSummary(ref, entry));
    return { items: all.slice(0, maxObjects), truncated: all.length > maxObjects };
  }

  objectTypeForRun(runId: number, objectIndex?: number): string | undefined {
    for (const entry of this.objects.values()) {
      if (entry.runId === runId && (objectIndex === undefined || entry.objectIndex === objectIndex)) return entry.type;
    }
    return undefined;
  }

  objectRefForRun(runId: number, objectIndex?: number): string | undefined {
    for (const [ref, entry] of this.objects) {
      if (entry.runId === runId && (objectIndex === undefined || entry.objectIndex === objectIndex)) return ref;
    }
    return undefined;
  }

  asset(ref: string): AgentAsset {
    const part = this.assetsByRef.get(ref);
    if (!part) throw new Error("Unknown asset reference");
    const bytes = this.doc.media(part);
    if (!bytes) throw new Error("Asset bytes are unavailable");
    return { ref, bytes, mediaType: mediaType(part) };
  }
}
