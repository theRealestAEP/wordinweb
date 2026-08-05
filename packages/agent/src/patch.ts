import { markerShape, type MarkerShape, type ProjectedLine } from "./project.js";
import type { AgentAnchorLine, AgentAnchorSegment, AgentOperation, AgentPatchHunk, AgentProjectionMode } from "./types.js";

/**
 * Translate projection hunks into the existing intent vocabulary.
 *
 * Every address a hunk produces is a stable ref plus a wire offset taken from
 * an anchor segment, so the agent never sees a wire offset and the engine
 * never guesses a rendered one. Structural work is expressed so that no
 * operation has to name a paragraph a previous operation created: merges run
 * before the text they concatenate, and splits run last, back to front, at
 * offsets inside a run that already exists.
 */

/** A marker change is a paragraph-property change, and the review model has no
 * tracked form for one. A suggested hunk refuses it instead of applying it
 * untracked behind the reviewer's back. */
function markerOperations(blockRef: string, current: MarkerShape, target: MarkerShape, suggest: boolean): AgentOperation[] {
  const operations: AgentOperation[] = [];
  if (current.heading !== target.heading) {
    if (suggest) throw new Error("A suggested hunk cannot change a heading marker. Send the structure change as a direct edit.");
    operations.push({ kind: "formatParagraph", blockRef, styleId: target.heading === undefined ? "Normal" : `Heading${target.heading}` });
  }
  if (current.list !== target.list) {
    if (suggest) throw new Error("A suggested hunk cannot change a list marker. Send the structure change as a direct edit.");
    operations.push({ kind: "setListType", blockRef, listKind: target.list ?? null });
  }
  return operations;
}

/** The wire address of a projection column. The first segment that covers the
 * column wins, so text typed at a run boundary joins the run before it — the
 * same run an editor caret would extend. */
function insertPosition(anchor: AgentAnchorLine, column: number): { runRef: string; offset: number } {
  for (const segment of anchor.segments) {
    if (segment.editable && segment.start <= column && column <= segment.end) {
      return { runRef: segment.runRef, offset: segment.wireStart + (column - segment.start) };
    }
  }
  throw new Error(`Projection line ${anchor.line} column ${column} is not inside editable text`);
}

interface TextEdit {
  operations: AgentOperation[];
  insert?: { runRef: string; offset: number; column: number };
}

function textEdit(source: ProjectedLine, targetBody: string, suggest: boolean, clampPrefix: number, allowSuffix: boolean): TextEdit {
  const anchor = source.anchor;
  const blockRef = anchor.blockRef!;
  const base = anchor.marker;
  const current = source.text.slice(base);

  let prefix = 0;
  while (prefix < current.length && prefix < targetBody.length && current[prefix] === targetBody[prefix]) prefix++;
  prefix = Math.min(prefix, clampPrefix);
  let suffix = 0;
  if (allowSuffix) {
    while (suffix < current.length - prefix && suffix < targetBody.length - prefix
      && current[current.length - 1 - suffix] === targetBody[targetBody.length - 1 - suffix]) suffix++;
  }
  const start = base + prefix;
  const end = base + current.length - suffix;
  const inserted = targetBody.slice(prefix, targetBody.length - suffix);

  const operations: AgentOperation[] = [];
  const ranges: Array<Record<string, unknown>> = [];
  const touched = anchor.segments.filter((segment) => segment.end > start && segment.start < end);
  for (const segment of touched) {
    if (!segment.editable) throw new Error(`Projection line ${anchor.line} cannot be rewritten across a non-text atom`);
  }
  // Right to left, so each removal leaves the offsets of the ones before it
  // untouched. Each range stays inside one w:t, which is the widest span the
  // delete intent accepts.
  for (let index = touched.length - 1; index >= 0; index--) {
    const segment = touched[index];
    const from = Math.max(start, segment.start) - segment.start;
    const to = Math.min(end, segment.end) - segment.start;
    if (to <= from) continue;
    const range = { blockRef, runRef: segment.runRef, start: segment.wireStart + from, end: segment.wireStart + to };
    if (suggest) ranges.unshift(range);
    else operations.push({ kind: "deleteText", ...range });
  }
  // A tracked deletion wraps the struck text in w:del and splits the run
  // around it, so the run this hunk would insert into no longer exists by the
  // time the insertion runs. One or the other applies cleanly; both in one
  // transaction do not.
  if (ranges.length > 0 && inserted.length > 0) {
    throw new Error("A suggested hunk can add text or remove text, not both. Send the removal and the addition as separate patches.");
  }
  if (ranges.length > 0) operations.push({ kind: "suggestRevision", ranges });
  if (inserted.length === 0) return { operations };

  // The insertion point is the start of the removed span, which every removal
  // above left alone.
  const at = insertPosition(anchor, start);
  operations.push({
    kind: "insertText",
    at: { blockRef, runRef: at.runRef, offset: at.offset },
    text: inserted,
    ...(suggest ? { suggest: true } : {}),
  });
  return { operations, insert: { ...at, column: start } };
}

function paragraphOperations(source: ProjectedLine, targets: string[], mode: AgentProjectionMode, suggest: boolean): AgentOperation[] {
  const anchor = source.anchor;
  const blockRef = anchor.blockRef!;
  const shapes = targets.map((line) => markerShape(line, mode));
  for (const shape of shapes) {
    if (shape.width !== shapes[0].width || shape.heading !== shapes[0].heading || shape.list !== shapes[0].list) {
      throw new Error("Every line a paragraph splits into must repeat the first line's structural marker");
    }
  }
  const bodies = targets.map((line, index) => line.slice(shapes[index].width));
  const operations = markerOperations(blockRef, markerShape(source.text.slice(0, anchor.marker), mode), shapes[0], suggest);
  const body = bodies.join("");
  const current = source.text.slice(anchor.marker);

  if (targets.length === 1) {
    if (current !== body) operations.push(...textEdit(source, body, suggest, Number.POSITIVE_INFINITY, true).operations);
    return operations;
  }

  const boundaries: number[] = [];
  let at = 0;
  for (let index = 0; index < bodies.length - 1; index++) {
    at += bodies[index].length;
    boundaries.push(at);
  }

  if (current === body) {
    // A pure split rewrites nothing, so every run keeps its own formatting.
    // Back to front: a split only moves content that follows its point, so
    // the offsets of the earlier boundaries never shift.
    for (let index = boundaries.length - 1; index >= 0; index--) {
      const position = insertPosition(anchor, anchor.marker + boundaries[index]);
      operations.push({ kind: "splitParagraph", at: { blockRef, runRef: position.runRef, offset: position.offset }, ...(suggest ? { suggest: true } : {}) });
    }
    return operations;
  }

  // Keep the text before the first boundary and insert everything from there
  // as one span, so every boundary lands in one run at a known wire offset.
  const edit = textEdit(source, body, suggest, boundaries[0], false);
  operations.push(...edit.operations);
  const insert = edit.insert;
  if (!insert) throw new Error("The patch could not place the split text");
  for (let index = boundaries.length - 1; index >= 0; index--) {
    const offset = insert.offset + (anchor.marker + boundaries[index] - insert.column);
    operations.push({ kind: "splitParagraph", at: { blockRef, runRef: insert.runRef, offset }, ...(suggest ? { suggest: true } : {}) });
  }
  return operations;
}

/** A merge keeps every run's identity and its offsets inside that run, so the
 * paragraph the merge produces is exactly these lines' segments end to end. */
function mergedLine(lines: ProjectedLine[]): ProjectedLine {
  const first = lines[0];
  let text = first.text;
  const segments: AgentAnchorSegment[] = [...first.anchor.segments];
  for (const line of lines.slice(1)) {
    const shift = text.length - line.anchor.marker;
    for (const segment of line.anchor.segments) {
      segments.push({ ...segment, start: segment.start + shift, end: segment.end + shift });
    }
    text += line.text.slice(line.anchor.marker);
  }
  return { text, anchor: { ...first.anchor, segments } };
}

function hunkOperations(lines: ProjectedLine[], hunk: AgentPatchHunk, mode: AgentProjectionMode, suggest: boolean): AgentOperation[] {
  const slice = lines.slice(hunk.startLine - 1, hunk.endLine);
  for (const line of slice) {
    if (line.anchor.role !== "paragraph" || !line.anchor.blockRef) {
      throw new Error(`Projection line ${line.anchor.line} is ${line.anchor.role} content and cannot be patched`);
    }
    if (!line.anchor.editable) throw new Error(`Projection line ${line.anchor.line} has no editable text identity`);
  }
  const targets = hunk.newText.split("\n");
  const paired = Math.min(slice.length, targets.length);
  const operations: AgentOperation[] = [];
  for (let index = 0; index < paired - 1; index++) {
    operations.push(...paragraphOperations(slice[index], [targets[index]], mode, suggest));
  }
  // The last paired paragraph absorbs the remainder: it swallows the extra old
  // paragraphs, or it splits into the extra new ones.
  const extra = slice.slice(paired);
  // Removing a paragraph break is a paragraph-mark deletion, and this compiler
  // only knows the direct form of it. Refuse rather than merge untracked.
  if (suggest && extra.length > 0) {
    throw new Error("A suggested hunk cannot remove a paragraph break. Send the merge as a direct edit.");
  }
  for (const line of extra) operations.push({ kind: "mergeParagraph", blockRef: line.anchor.blockRef! });
  const last = extra.length > 0 ? mergedLine(slice.slice(paired - 1)) : slice[paired - 1];
  operations.push(...paragraphOperations(last, targets.slice(paired - 1), mode, suggest));
  return operations;
}

export function patchOperations(lines: ProjectedLine[], hunks: AgentPatchHunk[], mode: AgentProjectionMode, suggest: boolean): AgentOperation[] {
  if (!Array.isArray(hunks) || hunks.length === 0) throw new Error("At least one edit hunk is required");
  const ordered = hunks.map((hunk, index) => {
    if (!hunk || typeof hunk !== "object") throw new Error("Each edit must be an object");
    const { startLine, endLine, newText } = hunk;
    if (!Number.isInteger(startLine) || !Number.isInteger(endLine) || typeof newText !== "string") {
      throw new Error("Each edit needs integer startLine and endLine and a newText string");
    }
    if (startLine < 1 || endLine < startLine || endLine > lines.length) {
      throw new Error(`Edit ${index + 1} covers lines ${startLine}-${endLine}, outside the projected 1-${lines.length}`);
    }
    return hunk;
  }).sort((left, right) => right.startLine - left.startLine);
  for (let index = 1; index < ordered.length; index++) {
    if (ordered[index].endLine >= ordered[index - 1].startLine) throw new Error("Edit hunks must not overlap");
  }
  return ordered.flatMap((hunk) => hunkOperations(lines, hunk, mode, suggest));
}

/** Unified diff as a convenience over the edit list. Context and removed lines
 * are checked against the projection, so a diff written against a different
 * revision fails instead of landing in the wrong place. */
export function hunksFromUnifiedDiff(diff: string, lines: string[]): AgentPatchHunk[] {
  if (typeof diff !== "string" || diff.length === 0) throw new Error("The diff is empty");
  const rows = diff.split("\n");
  const hunks: AgentPatchHunk[] = [];
  let index = 0;
  while (index < rows.length) {
    const header = /^@@ -(\d+)(?:,(\d+))? \+\d+(?:,\d+)? @@/.exec(rows[index]);
    if (!header) {
      index++;
      continue;
    }
    const start = Number(header[1]);
    const count = header[2] === undefined ? 1 : Number(header[2]);
    if (start < 1) throw new Error("A diff hunk that inserts before the first line is not supported; rewrite line 1 instead");
    index++;
    const kept: string[] = [];
    let consumed = 0;
    while (index < rows.length && !rows[index].startsWith("@@")) {
      const row = rows[index];
      const marker = row.length === 0 ? " " : row[0];
      const text = row.slice(1);
      if (marker === "\\") {
        index++;
        continue;
      }
      if (marker !== " " && marker !== "-" && marker !== "+") break;
      if (marker !== "+") {
        if (consumed >= count) break;
        const line = lines[start - 1 + consumed];
        if (line !== text) throw new Error(`The diff expected "${text}" at line ${start + consumed} but the projection has "${line ?? ""}"`);
        consumed++;
      }
      if (marker !== "-") kept.push(text);
      index++;
    }
    if (consumed !== count) throw new Error(`A diff hunk declared ${count} original lines and supplied ${consumed}`);
    // A zero-length original range inserts after `start`; the edit list always
    // rewrites at least one line, so carry that line through unchanged.
    if (count === 0) {
      if (start > lines.length) throw new Error(`A diff hunk inserts after line ${start}, past the projected ${lines.length}`);
      kept.unshift(lines[start - 1]);
    }
    if (kept.length === 0) throw new Error("A diff hunk must leave at least one line; merge into the preceding line instead");
    hunks.push({ startLine: start, endLine: start + Math.max(1, count) - 1, newText: kept.join("\n") });
  }
  if (hunks.length === 0) throw new Error("The diff contains no @@ hunks");
  return hunks;
}
