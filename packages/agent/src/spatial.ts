import {
  layoutDocument,
  type Block,
  type DocxDocument,
  type PageItem,
  type Paragraph,
  type Run,
  type StableIds,
  type XmlElement,
} from "@wordinweb/core";
import { blockRef, runRef } from "./refs.js";
import type { AgentOverlap, AgentSpatialObject, AgentSpatialResult, SpatialBounds } from "./types.js";
import type { SemanticInspector } from "./inspect.js";

interface Point { x: number; y: number }
interface InternalObject extends AgentSpatialObject {
  polygon: Point[];
  paintOrder: number;
}

function runsOf(paragraph: Paragraph): Run[] {
  return paragraph.children.flatMap((child) => child.type === "run" ? [child] : child.runs);
}

function visitBlocks(blocks: Block[], visit: (block: Block) => void): void {
  for (const block of blocks) {
    visit(block);
    if (block.type === "table") for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks, visit);
  }
}

function sourceMaps(doc: DocxDocument, ids: StableIds): {
  drawingRuns: Map<XmlElement, { runId: number; objectIndex: number }>;
  mathBlocks: Map<XmlElement, number>;
  tables: Map<XmlElement, number>;
} {
  const drawingRuns = new Map<XmlElement, { runId: number; objectIndex: number }>();
  const mathBlocks = new Map<XmlElement, number>();
  const tables = new Map<XmlElement, number>();
  const stories: Block[][] = [
    ...doc.sections.map((section) => section.blocks),
    ...[...doc.headers.values()].map((story) => story.blocks),
    ...[...doc.footers.values()].map((story) => story.blocks),
    ...doc.footnotes.values(),
    ...doc.endnotes.values(),
  ];
  for (const blocks of stories) visitBlocks(blocks, (block) => {
    if (!block.src) return;
    if (block.type === "table") {
      const id = ids.idOf(block.src);
      if (id !== undefined) tables.set(block.src, id);
      return;
    }
    const blockId = ids.idOf(block.src);
    for (const run of runsOf(block)) {
      if (!run.src) continue;
      const runId = ids.idOf(run.src);
      if (runId === undefined) continue;
      run.content.forEach((content, objectIndex) => {
        const target = { runId, objectIndex };
        if (content.kind === "image" && content.srcDrawing) drawingRuns.set(content.srcDrawing, target);
        else if (content.kind === "drawing" && content.srcDrawing) drawingRuns.set(content.srcDrawing, target);
        else if (content.kind === "anchor") {
          const shape = content.shape;
          if (shape.type === "wordart" && shape.src) drawingRuns.set(shape.src, target);
          else if (shape.type === "line" && shape.src) drawingRuns.set(shape.src, target);
          else if ((shape.type === "image" || shape.type === "textbox" || shape.type === "art") && shape.srcDrawing) drawingRuns.set(shape.srcDrawing, target);
        } else if (content.kind === "math" && content.src && blockId !== undefined) {
          mathBlocks.set(content.src, blockId);
        }
      });
    }
  });
  return { drawingRuns, mathBlocks, tables };
}

/** Exhaustive classification. A new layout primitive must make a deliberate
 * choice here before the agent package builds. */
function itemRole(item: PageItem): "text" | "object" | "decoration" | "ui" {
  switch (item.kind) {
    case "text": return "text";
    case "image":
    case "chart":
    case "drawingHit":
    case "wordart":
    case "warptext": return "object";
    case "path":
    case "rect":
    case "edge": return "decoration";
    case "grip": return "ui";
    default: {
      const exhaustive: never = item;
      return exhaustive;
    }
  }
}

function itemBounds(item: PageItem): SpatialBounds | null {
  switch (item.kind) {
    case "text": return { x: item.x, y: item.glyphTop ?? item.lineTop, width: item.width, height: item.glyphBoxH ?? item.lineHeight };
    case "path":
    case "rect":
    case "image":
    case "chart":
    case "drawingHit":
    case "wordart":
    case "warptext": return { x: item.x, y: item.y, width: item.width, height: item.height };
    case "edge": return { x: Math.min(item.x1, item.x2), y: Math.min(item.y1, item.y2), width: Math.abs(item.x2 - item.x1), height: Math.abs(item.y2 - item.y1) };
    case "grip": {
      if (item.axis !== "move") return null;
      return { x: item.x, y: item.y1, width: (item.x2 ?? item.x) - item.x, height: item.y2 - item.y1 };
    }
    default: {
      const exhaustive: never = item;
      return exhaustive;
    }
  }
}

function layer(item: { behind?: boolean; front?: boolean }): AgentSpatialObject["layer"] {
  return item.behind ? "behind-text" : item.front ? "in-front-of-text" : "body";
}

function rotate(point: Point, degrees: number, origin: Point): Point {
  const radians = degrees * Math.PI / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = point.x - origin.x;
  const dy = point.y - origin.y;
  return { x: origin.x + dx * cos - dy * sin, y: origin.y + dx * sin + dy * cos };
}

function polygon(bounds: SpatialBounds, degrees = 0, origin?: Point): Point[] {
  const points = [
    { x: bounds.x, y: bounds.y },
    { x: bounds.x + bounds.width, y: bounds.y },
    { x: bounds.x + bounds.width, y: bounds.y + bounds.height },
    { x: bounds.x, y: bounds.y + bounds.height },
  ];
  if (!degrees) return points;
  const center = origin ?? { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 };
  return points.map((point) => rotate(point, degrees, center));
}

function signedArea(points: Point[]): number {
  let area = 0;
  for (let i = 0; i < points.length; i++) {
    const a = points[i];
    const b = points[(i + 1) % points.length];
    area += a.x * b.y - b.x * a.y;
  }
  return area / 2;
}

function intersection(subject: Point[], clip: Point[]): Point[] {
  let output = subject;
  const clockwise = signedArea(clip) < 0;
  const inside = (point: Point, a: Point, b: Point) => {
    const cross = (b.x - a.x) * (point.y - a.y) - (b.y - a.y) * (point.x - a.x);
    return clockwise ? cross <= 1e-8 : cross >= -1e-8;
  };
  const lineIntersection = (start: Point, end: Point, a: Point, b: Point): Point => {
    const x1 = start.x; const y1 = start.y; const x2 = end.x; const y2 = end.y;
    const x3 = a.x; const y3 = a.y; const x4 = b.x; const y4 = b.y;
    const denominator = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
    if (Math.abs(denominator) < 1e-12) return end;
    return {
      x: ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / denominator,
      y: ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / denominator,
    };
  };
  for (let i = 0; i < clip.length; i++) {
    const input = output;
    output = [];
    if (!input.length) break;
    const a = clip[i];
    const b = clip[(i + 1) % clip.length];
    let start = input[input.length - 1];
    for (const end of input) {
      if (inside(end, a, b)) {
        if (!inside(start, a, b)) output.push(lineIntersection(start, end, a, b));
        output.push(end);
      } else if (inside(start, a, b)) output.push(lineIntersection(start, end, a, b));
      start = end;
    }
  }
  return output;
}

function boundsOf(points: Point[]): SpatialBounds {
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
}

function polygonArea(points: Point[]): number {
  return Math.abs(signedArea(points));
}

function topObject(a: InternalObject, b: InternalObject): string | undefined {
  const rank = (value: AgentSpatialObject["layer"]) => value === "behind-text" ? -1 : value === "in-front-of-text" ? 1 : 0;
  const layerDifference = rank(a.layer) - rank(b.layer);
  if (layerDifference !== 0) return layerDifference > 0 ? a.ref : b.ref;
  if (a.zOrder !== undefined || b.zOrder !== undefined) {
    const difference = (a.zOrder ?? 0) - (b.zOrder ?? 0);
    if (difference !== 0) return difference > 0 ? a.ref : b.ref;
  }
  return a.paintOrder === b.paintOrder ? undefined : a.paintOrder > b.paintOrder ? a.ref : b.ref;
}

export function spatialInspect(
  doc: DocxDocument,
  ids: StableIds,
  revision: string,
  inspector: SemanticInspector,
  pageRange?: { start: number; count: number },
  includeOverlaps = true,
): AgentSpatialResult {
  const result = layoutDocument(doc);
  const start = pageRange?.start ?? 1;
  const count = pageRange?.count ?? Math.min(result.totalPages, 10);
  if (!Number.isInteger(start) || start < 1 || !Number.isInteger(count) || count < 1 || count > 100) throw new Error("Invalid page range");
  const wanted = result.pages.filter((page) => page.index >= start && page.index < start + count);
  const maps = sourceMaps(doc, ids);
  const objects: InternalObject[] = [];
  const occurrences = new Map<string, number>();
  const add = (semanticRef: string, editRef: string | undefined, type: string, page: number, bounds: SpatialBounds, item: { behind?: boolean; front?: boolean; z?: number; rotate?: { deg: number; ox: number; oy: number }; rotation?: number }, paintOrder: number): void => {
    const occurrenceKey = `${semanticRef}@${page}`;
    const occurrence = (occurrences.get(occurrenceKey) ?? 0) + 1;
    occurrences.set(occurrenceKey, occurrence);
    const ref = `spatial:${semanticRef}:${page}:${occurrence}`;
    const degrees = item.rotate?.deg ?? item.rotation ?? 0;
    const origin = item.rotate ? { x: bounds.x + item.rotate.ox, y: bounds.y + item.rotate.oy } : undefined;
    const poly = polygon(bounds, degrees, origin);
    objects.push({ ref, editRef, type, page, bounds: boundsOf(poly), rotation: degrees || undefined, layer: layer(item), zOrder: item.z, polygon: poly, paintOrder });
  };

  for (const page of wanted) {
    const seen = new Set<string>();
    page.items.forEach((item, paintOrder) => {
      itemRole(item);
      const bounds = itemBounds(item);
      if (!bounds) return;
      if (item.kind === "drawingHit") {
        const target = maps.drawingRuns.get(item.src);
        if (!target) return;
        const semanticRef = inspector.objectRefForRun(target.runId, target.objectIndex) ?? runRef(target.runId);
        const key = `${semanticRef}:${page.index}`;
        if (seen.has(key)) return;
        seen.add(key);
        add(semanticRef, semanticRef, inspector.objectTypeForRun(target.runId, target.objectIndex) ?? "drawing", page.index, bounds, item, paintOrder);
      } else if (item.kind === "image" && item.src) {
        const target = maps.drawingRuns.get(item.src);
        if (!target) return;
        const semanticRef = inspector.objectRefForRun(target.runId, target.objectIndex) ?? runRef(target.runId);
        const key = `${semanticRef}:${page.index}`;
        if (seen.has(key)) return;
        seen.add(key);
        add(semanticRef, semanticRef, inspector.objectTypeForRun(target.runId, target.objectIndex) ?? "image", page.index, bounds, item, paintOrder);
      } else if (item.kind === "wordart" && item.src) {
        const target = maps.drawingRuns.get(item.src);
        if (!target) return;
        const semanticRef = inspector.objectRefForRun(target.runId, target.objectIndex) ?? runRef(target.runId);
        const key = `${semanticRef}:${page.index}`;
        if (seen.has(key)) return;
        seen.add(key);
        add(semanticRef, semanticRef, "word-art", page.index, bounds, { ...item, rotation: item.rotation }, paintOrder);
      } else if (item.kind === "text" && item.mathSrc) {
        const blockId = maps.mathBlocks.get(item.mathSrc);
        if (blockId === undefined) return;
        const semanticRef = `math:${blockId}`;
        const existing = objects.find((object) => object.page === page.index && object.editRef === blockRef(blockId) && object.type === "math");
        if (existing) {
          const combined = {
            x: Math.min(existing.bounds.x, bounds.x),
            y: Math.min(existing.bounds.y, bounds.y),
            width: Math.max(existing.bounds.x + existing.bounds.width, bounds.x + bounds.width) - Math.min(existing.bounds.x, bounds.x),
            height: Math.max(existing.bounds.y + existing.bounds.height, bounds.y + bounds.height) - Math.min(existing.bounds.y, bounds.y),
          };
          existing.bounds = combined;
          existing.polygon = polygon(combined);
        } else add(semanticRef, blockRef(blockId), "math", page.index, bounds, item, paintOrder);
      } else if (item.kind === "grip" && item.axis === "move") {
        const tableId = maps.tables.get(item.tbl);
        if (tableId !== undefined) add(blockRef(tableId), blockRef(tableId), "table", page.index, bounds, {}, paintOrder);
      }
    });
  }

  const overlaps: AgentOverlap[] = [];
  if (includeOverlaps) {
    for (let i = 0; i < objects.length; i++) for (let j = i + 1; j < objects.length; j++) {
      const a = objects[i];
      const b = objects[j];
      if (a.page !== b.page) continue;
      const hit = intersection(a.polygon, b.polygon);
      const area = hit.length >= 3 ? polygonArea(hit) : 0;
      if (area <= 1e-6) continue;
      overlaps.push({ objects: [a.ref, b.ref], overlapBounds: boundsOf(hit), overlapArea: area, topObject: topObject(a, b) });
    }
  }

  return {
    revision,
    layout: { quality: typeof document === "undefined" ? "approximate" : "exact", profile: typeof document === "undefined" ? "headless-approx" : "browser-canvas" },
    totalPages: result.totalPages,
    pages: wanted.map((page) => ({ index: page.index, width: page.width, height: page.height })),
    objects: objects.map(({ polygon: _polygon, paintOrder: _paintOrder, ...object }) => object),
    overlaps,
  };
}
