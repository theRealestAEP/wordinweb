import {
  imageAltText,
  type Block,
  type DocxDocument,
  type Paragraph,
  type RunContent,
  type StableIds,
  wireOffsetOf,
} from "@wordinweb/core";
import { mathText, paragraphRuns, runText } from "./inspect.js";
import { blockRef, objectRef, runRef } from "./refs.js";
import type {
  AgentAnchorLine,
  AgentAnchorSegment,
  AgentProjectionMode,
  AgentProjectResult,
} from "./types.js";

/**
 * DocMD: a deterministic text projection of a story, produced in the same pass
 * as the anchor map that owns the rendered-to-wire offset translation.
 *
 * Inspection offsets are rendered text and edit offsets are wire units, and
 * reconciling the two used to be the caller's problem. Here the two spaces are
 * derived together: every projected character belongs to exactly one anchor
 * segment, and a segment is editable only when it came verbatim from one `w:t`
 * — the one case where a projection column is also a wire offset. Agents read
 * the text; the segments do the arithmetic.
 */

/** Single-character stand-ins for inline content with no editable text. One
 * character each, so a projection column always names one whole atom and the
 * anchor map can mark it opaque without span arithmetic. */
export const DOCMD_ATOMS = {
  tab: "␉",
  lineBreak: "␊",
  pageBreak: "␌",
  field: "␎",
  math: "␏",
  noteReference: "␅",
  object: "￼",
} as const;

export interface ProjectedLine {
  text: string;
  anchor: AgentAnchorLine;
}

function fieldName(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() || "FIELD";
}

function renderContent(content: RunContent, mode: AgentProjectionMode, runId: number | undefined, index: number): string {
  const object = (label: string): string =>
    mode === "text" || runId === undefined ? DOCMD_ATOMS.object : `![${label}](${objectRef(runId, index)})`;
  switch (content.kind) {
    // A literal newline inside a w:t would add a projection line the anchor
    // map does not know about; the stand-in keeps the mapping one-to-one.
    case "text": return content.text.replace(/[\r\n]/g, DOCMD_ATOMS.lineBreak);
    case "break": return content.breakType === "line" ? DOCMD_ATOMS.lineBreak : DOCMD_ATOMS.pageBreak;
    case "tab":
    case "ptab": return DOCMD_ATOMS.tab;
    case "field": return mode === "text" ? DOCMD_ATOMS.field : `{{${fieldName(content.instruction)}}}`;
    case "math": return mode === "text" ? DOCMD_ATOMS.math : `$${mathText(content.nodes)}$`;
    case "noteRef": return mode === "text" ? DOCMD_ATOMS.noteReference : `[^${content.id}]`;
    case "ruby": return runText(content.base);
    case "image": return object(content.srcDrawing ? imageAltText(content.srcDrawing) : "");
    case "anchor": return object(content.shape.type);
    case "drawing": return object(content.chart ? "chart" : content.smartArt ? "smart-art" : content.textbox ? "text-box" : "drawing");
    default: {
      const exhaustive: never = content;
      return exhaustive;
    }
  }
}

const HEADING_MARKER = /^(#{1,6}) /;
const LIST_MARKER = /^ *(?:[-*+]|\d+[.)]) /;

export interface MarkerShape {
  width: number;
  heading?: number;
  list?: "bullet" | "number";
}

/** Read the structural prefix of a projection line. The projector writes the
 * only prefixes this has to recognize, and it escapes paragraph text that
 * would otherwise read as one, so the parse is never a guess. */
export function markerShape(line: string, mode: AgentProjectionMode): MarkerShape {
  if (mode === "text") return { width: 0 };
  if (line.startsWith("\\")) return { width: 1 };
  const heading = HEADING_MARKER.exec(line);
  if (heading) return { width: heading[0].length, heading: heading[1].length };
  const list = LIST_MARKER.exec(line);
  if (list) return { width: list[0].length, list: /[-*+] $/.test(list[0]) ? "bullet" : "number" };
  return { width: 0 };
}

function paragraphMarker(doc: DocxDocument, paragraph: Paragraph, mode: AgentProjectionMode, body: string): string {
  if (mode === "text") return "";
  const effective = doc.effectiveParaProps(paragraph);
  if (effective.outlineLevel !== undefined) return `${"#".repeat(Math.min(6, effective.outlineLevel + 1))} `;
  const numbering = effective.numbering;
  if (numbering) {
    const level = doc.numberingLevel(numbering.numId, numbering.ilvl);
    return `${"  ".repeat(Math.min(8, numbering.ilvl))}${level?.format === "bullet" ? "- " : "1. "}`;
  }
  // Plain paragraphs really do open with "3. " or "- ". Without the escape a
  // patch would read the document's own text as a list marker and strip it.
  return markerShape(body, mode).width > 0 ? "\\" : "";
}

export function projectParagraph(doc: DocxDocument, ids: StableIds, paragraph: Paragraph, mode: AgentProjectionMode): ProjectedLine {
  const blockId = paragraph.src ? ids.idOf(paragraph.src) : undefined;
  const segments: AgentAnchorSegment[] = [];
  let body = "";
  for (const { run } of paragraphRuns(paragraph)) {
    const runId = run.src ? ids.idOf(run.src) : undefined;
    run.content.forEach((content, index) => {
      const rendered = renderContent(content, mode, runId, index);
      const wire = runId !== undefined && run.src && content.kind === "text" && content.srcT
        ? wireOffsetOf(run.src, content.srcT, 0)
        : null;
      // An empty w:t still earns a zero-width segment: it is where text goes
      // when an agent types into an empty paragraph.
      if (rendered.length === 0 && wire === null) return;
      const start = body.length;
      body += rendered;
      segments.push({
        start,
        end: body.length,
        runRef: runId === undefined ? "" : runRef(runId),
        wireStart: wire ?? 0,
        wireEnd: wire === null ? 0 : wire + rendered.length,
        editable: wire !== null,
      });
    });
  }
  const marker = paragraphMarker(doc, paragraph, mode, body);
  for (const segment of segments) {
    segment.start += marker.length;
    segment.end += marker.length;
  }
  return {
    text: marker + body,
    anchor: {
      line: 0,
      role: "paragraph",
      ...(blockId === undefined ? {} : { blockRef: blockRef(blockId) }),
      marker: marker.length,
      editable: blockId !== undefined && segments.some((segment) => segment.editable),
      segments,
    },
  };
}

/** The stories a projection can address, matching the ids `overview` reports.
 * Content inside drawings stays reachable through `object` inspection. */
export function projectionStories(doc: DocxDocument): Map<string, Block[]> {
  const stories = new Map<string, Block[]>();
  stories.set("body", doc.sections.flatMap((section) => section.blocks));
  for (const [id, value] of doc.headers) stories.set(`header:${id}`, value.blocks);
  for (const [id, value] of doc.footers) stories.set(`footer:${id}`, value.blocks);
  for (const [id, blocks] of doc.footnotes) stories.set(`footnote:${id}`, blocks);
  for (const [id, blocks] of doc.endnotes) stories.set(`endnote:${id}`, blocks);
  return stories;
}

export function projectStory(
  doc: DocxDocument,
  ids: StableIds,
  revision: string,
  storyId = "body",
  mode: AgentProjectionMode = "md",
  cursorValue?: string,
  maxBlocks = 200,
  maxCharacters = 24_000,
): AgentProjectResult {
  if (mode !== "text" && mode !== "md" && mode !== "outline") throw new Error("mode must be text, md, or outline");
  if (!Number.isInteger(maxBlocks) || maxBlocks < 1 || maxBlocks > 2000) throw new Error("maxBlocks must be between 1 and 2000");
  if (!Number.isInteger(maxCharacters) || maxCharacters < 1 || maxCharacters > 200_000) throw new Error("maxCharacters must be between 1 and 200000");
  const blocks = projectionStories(doc).get(storyId);
  if (!blocks) throw new Error("Unknown story");

  let index = 0;
  if (cursorValue !== undefined) {
    const match = /^docmd:(.*):(\d+)$/.exec(cursorValue);
    if (!match || match[1] !== revision) throw new Error("Invalid or stale cursor");
    index = Number(match[2]);
  }

  const lines: string[] = [];
  const anchors: AgentAnchorLine[] = [];
  const emit = (text: string, anchor: AgentAnchorLine): void => {
    lines.push(text);
    anchors.push({ ...anchor, line: lines.length });
  };
  const structure: AgentAnchorLine = { line: 0, role: "structure", marker: 0, editable: false, segments: [] };

  // A GFM cell is one line, so a cell's paragraphs — and any table nested
  // inside it — flatten into one run of text rather than disappearing.
  const cellParts = (cellBlocks: Block[]): string[] => cellBlocks.flatMap((block) => {
    if (block.type === "table") return block.rows.flatMap((row) => row.cells.flatMap((cell) => cellParts(cell.blocks)));
    const line = projectParagraph(doc, ids, block, mode);
    return [line.text.slice(line.anchor.marker)];
  });

  const cellText = (cellBlocks: Block[]): string => cellParts(cellBlocks)
    .filter((value) => value.length > 0)
    .join(" ")
    .replace(/\|/g, "\\|");

  const emitTable = (table: Extract<Block, { type: "table" }>): void => {
    const tableId = table.src ? ids.idOf(table.src) : undefined;
    const anchor: AgentAnchorLine = {
      line: 0,
      role: "table",
      ...(tableId === undefined ? {} : { blockRef: blockRef(tableId) }),
      marker: 0,
      editable: false,
      segments: [],
    };
    const columns = Math.max(1, ...table.rows.map((row) => row.cells.length));
    emit("", structure);
    table.rows.forEach((row, rowIndex) => {
      emit(`| ${Array.from({ length: columns }, (_, column) => cellText(row.cells[column]?.blocks ?? [])).join(" | ")} |`, anchor);
      if (rowIndex === 0) emit(`| ${Array.from({ length: columns }, () => "---").join(" | ")} |`, anchor);
    });
    emit("", structure);
  };

  const emitBlock = (block: Block): void => {
    if (block.type === "paragraph") {
      if (mode === "outline" && doc.effectiveParaProps(block).outlineLevel === undefined) return;
      const projected = projectParagraph(doc, ids, block, mode);
      emit(projected.text, projected.anchor);
      return;
    }
    if (mode === "outline") return;
    // text mode is the document as a text file: a cell paragraph is a
    // paragraph and gets its own editable line. md mode owes the reader GFM.
    if (mode === "text") {
      for (const row of block.rows) for (const cell of row.cells) for (const inner of cell.blocks) emitBlock(inner);
      return;
    }
    emitTable(block);
  };

  const startBlock = index;
  let used = 0;
  while (index < blocks.length) {
    if (lines.length > 0 && (index - startBlock >= maxBlocks || used >= maxCharacters)) break;
    const before = lines.length;
    emitBlock(blocks[index]);
    for (let line = before; line < lines.length; line++) used += lines[line].length + 1;
    index++;
  }

  const truncated = index < blocks.length;
  return {
    revision,
    story: storyId,
    mode,
    text: lines.join("\n"),
    lines: lines.length,
    window: { cursor: cursorValue ?? null, startBlock, endBlock: index, maxBlocks, maxCharacters },
    anchors,
    ...(truncated ? { next: { value: `docmd:${revision}:${index}` } } : {}),
    truncated,
  };
}

/** Pair each projected line with its anchor. No projected character is a
 * newline, so the split recovers exactly the lines the projection emitted. */
export function projectedLines(projection: AgentProjectResult): ProjectedLine[] {
  const texts = projection.text.length === 0 && projection.lines === 0 ? [] : projection.text.split("\n");
  if (texts.length !== projection.anchors.length) throw new Error("The projection text and anchor map disagree");
  return texts.map((text, index) => ({ text, anchor: projection.anchors[index] }));
}
