import { DocxDocument } from "../docx.js";
import { Block, Paragraph, Run } from "../model.js";
import { XmlElement } from "../xml.js";
import { FormattedRange, SelectionSegment } from "./commands.js";

/**
 * Find & replace and selection text transforms. Matching runs over the
 * parsed model (body text incl. tables); positions map back to source w:t
 * elements so hits can be selected, replaced, or case-transformed.
 */

export interface FindMatch {
  /** Covered source ranges, in order (a match may span several runs). */
  ranges: FormattedRange[];
}

interface CharRef {
  t: XmlElement;
  offset: number;
}

function paragraphsIn(doc: DocxDocument): Paragraph[] {
  const out: Paragraph[] = [];
  const fromBlocks = (blocks: readonly Block[]): void => {
    for (const b of blocks) {
      if (b.type === "paragraph") out.push(b);
      else if (b.type === "table") {
        for (const row of b.rows) for (const cell of row.cells) fromBlocks(cell.blocks);
      }
    }
  };
  for (const s of doc.sections) fromBlocks(s.blocks);
  return out;
}

function runsOf(para: Paragraph): Run[] {
  const out: Run[] = [];
  for (const c of para.children) {
    if (c.type === "run") out.push(c);
    else out.push(...c.runs);
  }
  return out;
}

/** All matches of `query` in body text (paragraph-local, no cross-para). */
export function findAll(doc: DocxDocument, query: string, opts?: { matchCase?: boolean }): FindMatch[] {
  if (!query) return [];
  const norm = (s: string) => (opts?.matchCase ? s : s.toLowerCase());
  const q = norm(query);
  const matches: FindMatch[] = [];

  for (const para of paragraphsIn(doc)) {
    let text = "";
    const refs: CharRef[] = [];
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
    const hay = norm(text);
    let from = 0;
    for (;;) {
      const idx = hay.indexOf(q, from);
      if (idx === -1) break;
      from = idx + q.length;
      // Convert char span -> per-t ranges.
      const ranges: FormattedRange[] = [];
      for (let i = idx; i < idx + q.length; i++) {
        const ref = refs[i];
        const last = ranges[ranges.length - 1];
        if (last && last.t === ref.t && last.end === ref.offset) last.end = ref.offset + 1;
        else ranges.push({ t: ref.t, start: ref.offset, end: ref.offset + 1 });
      }
      matches.push({ ranges });
    }
  }
  return matches;
}

/**
 * Replace one match with `replacement`: the first covered range takes the
 * new text, the rest of the match is deleted. Returns the resulting range.
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

/** Replace every occurrence; returns the number of replacements. */
export function replaceAll(doc: DocxDocument, query: string, replacement: string, opts?: { matchCase?: boolean }): number {
  // One find pass, applied back-to-front: replacing at later offsets never
  // shifts earlier match positions (and re-containing replacements can't
  // loop, since the pass is fixed up front).
  const matches = findAll(doc, query, opts);
  for (let i = matches.length - 1; i >= 0; i--) replaceMatch(doc, matches[i], replacement);
  return matches.length;
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
