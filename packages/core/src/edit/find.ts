import { DocxDocument } from "../docx.js";
import { Block, Paragraph, Run } from "../model.js";
import { XmlElement } from "../xml.js";
import { FormattedRange, SelectionSegment } from "./commands.js";
import { EncodedCaret } from "./ids.js";

/**
 * Find & replace and selection text transforms. Matching runs over the
 * parsed model — every story: the body (incl. tables), each header and
 * footer part, and each footnote and endnote. Positions map back to source
 * w:t elements so hits can be selected, replaced, or case-transformed.
 *
 * Within one story, paragraphs concatenate with "\n" standing for the
 * paragraph mark (Word's ^p), so a query containing "\n" matches across
 * paragraph boundaries. A match never crosses a story boundary.
 */

/** Which story a match sits in. Every header part is "header" (and every
 * footer part "footer") regardless of which section references it. */
export type FindStory = "body" | "header" | "footer" | "footnote" | "endnote";

export interface FindOptions {
  /** Case-sensitive matching. Default: case-insensitive. */
  matchCase?: boolean;
  /** Match whole words only: the characters adjacent to the hit must not be
   * letters, digits, or underscores (Word's "Find whole words only"). */
  wholeWord?: boolean;
}

export interface FindMatch {
  /** Covered source ranges, in order (a match may span several runs, and —
   * when the query contains "\n" — several paragraphs). */
  ranges: FormattedRange[];
  story: FindStory;
}

interface CharRef {
  t: XmlElement;
  offset: number;
}

function paragraphsInBlocks(blocks: readonly Block[]): Paragraph[] {
  const out: Paragraph[] = [];
  const fromBlocks = (list: readonly Block[]): void => {
    for (const b of list) {
      if (b.type === "paragraph") out.push(b);
      else if (b.type === "table") {
        for (const row of b.rows) for (const cell of row.cells) fromBlocks(cell.blocks);
      }
    }
  };
  fromBlocks(blocks);
  return out;
}

/** One searchable text sequence. The body is one unit; each header, footer,
 * footnote, and endnote is its own (matches never span parts). */
function storyUnits(doc: DocxDocument): { story: FindStory; paragraphs: Paragraph[] }[] {
  const units: { story: FindStory; paragraphs: Paragraph[] }[] = [];
  const body: Paragraph[] = [];
  for (const s of doc.sections) body.push(...paragraphsInBlocks(s.blocks));
  units.push({ story: "body", paragraphs: body });
  for (const hf of doc.headers.values()) units.push({ story: "header", paragraphs: paragraphsInBlocks(hf.blocks) });
  for (const hf of doc.footers.values()) units.push({ story: "footer", paragraphs: paragraphsInBlocks(hf.blocks) });
  for (const blocks of doc.footnotes.values()) units.push({ story: "footnote", paragraphs: paragraphsInBlocks(blocks) });
  for (const blocks of doc.endnotes.values()) units.push({ story: "endnote", paragraphs: paragraphsInBlocks(blocks) });
  return units;
}

function runsOf(para: Paragraph): Run[] {
  const out: Run[] = [];
  for (const c of para.children) {
    if (c.type === "run") out.push(c);
    else out.push(...c.runs);
  }
  return out;
}

const isWordChar = (ch: string | undefined): boolean =>
  ch !== undefined && /[\p{L}\p{N}_]/u.test(ch);

/** All matches of `query`, body first, then headers/footers/notes. */
export function findAll(doc: DocxDocument, query: string, opts?: FindOptions): FindMatch[] {
  if (!query) return [];
  const norm = (s: string) => (opts?.matchCase ? s : s.toLowerCase());
  const q = norm(query);
  const matches: FindMatch[] = [];

  for (const unit of storyUnits(doc)) {
    let text = "";
    // One entry per character of `text`; null at a "\n" paragraph mark,
    // which has no w:t to select or replace.
    const refs: (CharRef | null)[] = [];
    for (const para of unit.paragraphs) {
      if (refs.length > 0) {
        text += "\n";
        refs.push(null);
      }
      for (const run of runsOf(para)) {
        for (const c of run.content) {
          if (c.kind !== "text" || !c.srcT) continue;
          const t = c.srcT as XmlElement;
          for (let i = 0; i < c.text.length; i++) {
            text += c.text[i];
            refs.push({ t, offset: i });
          }
        }
      }
    }
    const hay = norm(text);
    let from = 0;
    for (;;) {
      const idx = hay.indexOf(q, from);
      if (idx === -1) break;
      from = idx + q.length;
      if (opts?.wholeWord && (isWordChar(text[idx - 1]) || isWordChar(text[idx + q.length]))) continue;
      // Convert char span -> per-t ranges (paragraph marks have none).
      const ranges: FormattedRange[] = [];
      for (let i = idx; i < idx + q.length; i++) {
        const ref = refs[i];
        if (!ref) continue;
        const last = ranges[ranges.length - 1];
        if (last && last.t === ref.t && last.end === ref.offset) last.end = ref.offset + 1;
        else ranges.push({ t: ref.t, start: ref.offset, end: ref.offset + 1 });
      }
      if (ranges.length > 0) matches.push({ ranges, story: unit.story });
    }
  }
  return matches;
}

/**
 * Replace one match with `replacement`: the first covered range takes the
 * new text, the rest of the match is deleted. A match that spans paragraphs
 * keeps its paragraph marks (only text is replaced). Returns the resulting
 * range.
 */
export function replaceMatch(doc: DocxDocument, match: FindMatch, replacement: string): FormattedRange | null {
  const first = match.ranges[0];
  if (!first) return null;
  for (let i = match.ranges.length - 1; i >= 1; i--) {
    const r = match.ranges[i];
    r.t.text = r.t.text.slice(0, r.start) + r.t.text.slice(r.end);
  }
  first.t.text = first.t.text.slice(0, first.start) + replacement + first.t.text.slice(first.end);
  doc.refresh();
  return { t: first.t, start: first.start, end: first.start + replacement.length };
}

export interface ReplaceAllResult {
  total: number;
  byStory: Partial<Record<FindStory, number>>;
}

/** Replace every occurrence; reports the count, total and per story. */
export function replaceAll(
  doc: DocxDocument,
  query: string,
  replacement: string,
  opts?: FindOptions,
): ReplaceAllResult {
  // One find pass, applied back-to-front: replacing at later offsets never
  // shifts earlier match positions (and re-containing replacements can't
  // loop, since the pass is fixed up front). Stories share no w:t elements,
  // so ordering across them is free.
  const matches = findAll(doc, query, opts);
  for (let i = matches.length - 1; i >= 0; i--) replaceMatch(doc, matches[i], replacement);
  const byStory: Partial<Record<FindStory, number>> = {};
  for (const m of matches) byStory[m.story] = (byStory[m.story] ?? 0) + 1;
  return { total: matches.length, byStory };
}

/** A wire range inside one run: the shape suggestRevision and presence use. */
export interface WireRange {
  blockId: number;
  runId: number;
  start: number;
  end: number;
}

/** Intent bodies a compiled replace produces (IntentBase fields are the
 * submitting client's to add). */
export type ReplaceIntentBody =
  | { kind: "deleteText"; blockId: number; runId: number; start: number; end: number }
  | { kind: "insertText"; at: { blockId: number; runId: number; offset: number }; text: string; suggest?: { author: string; date: string } }
  | { kind: "suggestRevision"; ranges: WireRange[]; suggest: { author: string; date: string } };

export interface ReplaceAllCompilation {
  intents: ReplaceIntentBody[];
  result: ReplaceAllResult;
}

/**
 * Compile one match into wire intents against the CURRENT tree, in the order
 * they must apply. Plain: per-range deleteText, rightmost first, then the
 * replacement inserted at the match's left edge. Suggesting: one
 * suggestRevision striking every range, then a tracked insertText at the left
 * edge — the patch machinery's strike-then-insert shape (a strike keeps the
 * addressed run in place holding everything before it, so the insertion
 * still resolves). Null when any range is not id-addressable (math
 * internals): in a room an unaddressable replace must be an honest no-op.
 */
export function compileReplaceMatch(
  doc: DocxDocument,
  match: FindMatch,
  replacement: string,
  suggest?: { author: string; date: string },
): ReplaceIntentBody[] | null {
  const ids = doc.stableIds;
  if (!ids || match.ranges.length === 0) return null;
  const wire: WireRange[] = [];
  for (const r of match.ranges) {
    const enc: EncodedCaret | null = ids.encodeCaret(r.t, r.start, (el) => doc.findParentOf(el) ?? null);
    if (!enc) return null;
    wire.push({ blockId: enc.blockId, runId: enc.runId, start: enc.offset, end: enc.offset + (r.end - r.start) });
  }
  const first = wire[0];
  const intents: ReplaceIntentBody[] = [];
  if (suggest) {
    intents.push({ kind: "suggestRevision", ranges: wire, suggest });
    if (replacement.length > 0) {
      intents.push({ kind: "insertText", at: { blockId: first.blockId, runId: first.runId, offset: first.start }, text: replacement, suggest });
    }
    return intents;
  }
  for (let i = wire.length - 1; i >= 0; i--) {
    const r = wire[i];
    intents.push({ kind: "deleteText", blockId: r.blockId, runId: r.runId, start: r.start, end: r.end });
  }
  if (replacement.length > 0) {
    intents.push({ kind: "insertText", at: { blockId: first.blockId, runId: first.runId, offset: first.start }, text: replacement });
  }
  return intents;
}

/**
 * Compile a whole replace-all into wire intents (collab's replaceAll). One
 * find pass fixed up front, compiled back-to-front: every intent's offsets
 * are encoded against the pristine tree, and processing later matches first
 * keeps them valid — a mutation at offset p never moves text before p, and
 * both suggest mutations keep the addressed run (and its truncated-in-place
 * w:t) holding everything before the edit. Unaddressable matches are skipped
 * and excluded from the counts.
 */
export function compileReplaceAll(
  doc: DocxDocument,
  query: string,
  replacement: string,
  opts?: FindOptions,
  suggest?: { author: string; date: string },
): ReplaceAllCompilation {
  const matches = findAll(doc, query, opts);
  const intents: ReplaceIntentBody[] = [];
  const byStory: Partial<Record<FindStory, number>> = {};
  let total = 0;
  for (let i = matches.length - 1; i >= 0; i--) {
    const compiled = compileReplaceMatch(doc, matches[i], replacement, suggest);
    if (!compiled) continue;
    intents.push(...compiled);
    total++;
    byStory[matches[i].story] = (byStory[matches[i].story] ?? 0) + 1;
  }
  return { intents, result: { total, byStory } };
}

/** Change the case of the selected text (mutates w:t text in place). */
export function transformCase(
  doc: DocxDocument,
  segments: SelectionSegment[],
  mode: "upper" | "lower" | "title",
): FormattedRange[] {
  const apply = (s: string): string =>
    mode === "upper"
      ? s.toUpperCase()
      : mode === "lower"
        ? s.toLowerCase()
        : s.replace(/\p{L}[\p{L}\p{M}'’]*/gu, (word) => word[0].toUpperCase() + word.slice(1).toLowerCase());
  const out: FormattedRange[] = [];
  for (const seg of segments) {
    if (!seg.t) continue;
    const t = seg.t;
    const start = Math.max(0, seg.start);
    const end = Math.min(t.text.length, seg.end);
    if (end <= start) continue;
    t.text = t.text.slice(0, start) + apply(t.text.slice(start, end)) + t.text.slice(end);
    out.push({ t, start, end });
  }
  if (out.length > 0) doc.refresh();
  return out;
}
