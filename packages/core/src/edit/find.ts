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
  /** Word's wildcard pattern language (the subset below). Like Word, wildcard
   * matching is always case-sensitive and ignores wholeWord — the pattern's
   * own `<` and `>` express word boundaries. */
  wildcards?: boolean;
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

// ---------------------------------------------------------------------------
// Query compilation: special characters and wildcards
// ---------------------------------------------------------------------------

/**
 * Longest query either mode compiles. The cap bounds what a pattern can cost:
 * a wildcard query becomes a regular expression run over every story, and an
 * unbounded pattern is an unbounded amount of backtracking to hand an
 * arbitrary caller.
 */
export const MAX_FIND_PATTERN = 256;

/** Most `*` / `@` quantifiers one wildcard pattern may hold. Each quantifier
 * multiplies the regex engine's worst-case backtracking, so the count is
 * capped rather than trusted (input-length caps alone do not bound it). */
const MAX_WILDCARDS = 8;

const escapeRegex = (ch: string): string => ch.replace(/[.*+?^${}()|[\]\\/-]/g, "\\$&");

/** The word-boundary classes Word's `<` and `>` (and wholeWord) use — the
 * same rule as isWordChar above. */
const NOT_BEFORE_WORD = "(?<![\\p{L}\\p{N}_])";
const NOT_AFTER_WORD = "(?![\\p{L}\\p{N}_])";

/**
 * Translate the LITERAL mode's caret escapes (Word's "special characters").
 * The modeled subset, against the story text this module builds:
 *
 *   ^p  paragraph mark ("\n" join)   ^t  tab            ^l  line break (w:br)
 *   ^#  any digit                    ^$  any letter     ^?  any character
 *   ^w  white space (a run of it)    ^s  nonbreaking space
 *   ^~  nonbreaking hyphen (U+2011)  ^-  optional hyphen (w:softHyphen)
 *   ^^  a caret
 *
 * An unknown escape stays literal text, so a stray caret finds itself. Null
 * when the query needs no translation — the caller keeps the plain
 * substring path.
 */
function literalPattern(query: string): string | null {
  if (!query.includes("^")) return null;
  let out = "";
  for (let i = 0; i < query.length; i++) {
    const ch = query[i];
    if (ch !== "^" || i + 1 >= query.length) {
      out += escapeRegex(ch);
      continue;
    }
    const code = query[++i];
    switch (code) {
      case "p": out += "\\n"; break;
      case "t": out += "\\t"; break;
      case "l": out += "\\v"; break;
      case "#": out += "[0-9]"; break;
      case "$": out += "\\p{L}"; break;
      case "?": out += "[^\\n]"; break;
      case "w": out += "[ \\t\\u00A0]+"; break;
      case "s": out += "\\u00A0"; break;
      case "~": out += "\\u2011"; break;
      case "-": out += "\\u00AD"; break;
      case "^": out += "\\^"; break;
      default: out += `\\^${escapeRegex(code)}`;
    }
  }
  return out;
}

/**
 * Translate Word's WILDCARD pattern language — the documented subset:
 *
 *   ?         any single character            *   any run of characters (lazy)
 *   [abc]     a character in the set          [!abc]  one not in the set
 *   [a-z]     a range (sets may mix both)     @   one or more of the previous
 *   <  start of a word                        >   end of a word
 *   \x        the literal character x         ^13 paragraph mark   ^9 tab
 *
 * NOT modeled (Word has them; a pattern using one stays literal or is
 * refused): `{n,m}` counts, `(…)` groups with `\n` back-references in the
 * replacement, and the `^0nnn` character codes beyond ^13/^9. `?`, `*` and
 * sets never cross a paragraph mark — Word's wildcards do not either.
 *
 * Null when the pattern is malformed (unterminated set, `@` with nothing
 * before it, too many quantifiers): the caller reports zero matches rather
 * than guessing.
 */
function wildcardPattern(query: string): string | null {
  let out = "";
  /** The last single unit emitted, for `@` to quantify; "" when the last
   * token cannot be quantified. */
  let last = "";
  let quantifiers = 0;
  const emit = (unit: string) => {
    out += unit;
    last = unit;
  };
  for (let i = 0; i < query.length; i++) {
    const ch = query[i];
    if (ch === "?") { emit("[^\\n]"); continue; }
    if (ch === "*") {
      if (++quantifiers > MAX_WILDCARDS) return null;
      out += "[^\\n]*?";
      last = "";
      continue;
    }
    if (ch === "@") {
      if (last === "") return null;
      if (++quantifiers > MAX_WILDCARDS) return null;
      out = out.slice(0, out.length - last.length) + `(?:${last})+`;
      last = "";
      continue;
    }
    if (ch === "<") { out += NOT_BEFORE_WORD; last = ""; continue; }
    if (ch === ">") { out += NOT_AFTER_WORD; last = ""; continue; }
    if (ch === "\\") {
      if (i + 1 >= query.length) return null;
      emit(escapeRegex(query[++i]));
      continue;
    }
    if (ch === "^") {
      const m = /^\^(13|9)/.exec(query.slice(i));
      if (!m) return null; // other ^nnn codes are not modeled
      emit(m[1] === "13" ? "\\n" : "\\t");
      i += m[1].length;
      continue;
    }
    if (ch === "[") {
      const negated = query[i + 1] === "!";
      let j = i + (negated ? 2 : 1);
      let body = "";
      for (; j < query.length && query[j] !== "]"; j++) {
        // A range keeps its "-" when it sits between two set members.
        if (query[j] === "-" && body.length > 0 && j + 1 < query.length && query[j + 1] !== "]") body += "-";
        else body += escapeRegex(query[j]);
      }
      if (j >= query.length || body.length === 0) return null; // unterminated or empty
      emit(`[${negated ? "^" : ""}${body}]`);
      i = j;
      continue;
    }
    emit(escapeRegex(ch));
  }
  return out.length > 0 ? out : null;
}

/** A compiled query: the plain-substring fast path, a regular expression, or
 * null for a query with no possible matches (empty, over the cap, or a
 * malformed wildcard pattern). */
type CompiledQuery = { literal: string } | { re: RegExp } | null;

export function compileFindQuery(query: string, opts?: FindOptions): CompiledQuery {
  if (!query || query.length > MAX_FIND_PATTERN) return null;
  if (opts?.wildcards) {
    const source = wildcardPattern(query);
    if (source === null) return null;
    try {
      // Always case-sensitive, Word's wildcard rule.
      return { re: new RegExp(source, "gu") };
    } catch {
      return null;
    }
  }
  const source = literalPattern(query);
  if (source === null) return { literal: opts?.matchCase ? query : query.toLowerCase() };
  const wrapped = opts?.wholeWord ? `${NOT_BEFORE_WORD}(?:${source})${NOT_AFTER_WORD}` : source;
  try {
    return { re: new RegExp(wrapped, opts?.matchCase ? "gu" : "giu") };
  } catch {
    return null;
  }
}

/** All matches of `query`, body first, then headers/footers/notes. */
export function findAll(doc: DocxDocument, query: string, opts?: FindOptions): FindMatch[] {
  const compiled = compileFindQuery(query, opts);
  if (!compiled) return [];
  const matches: FindMatch[] = [];

  for (const unit of storyUnits(doc)) {
    let text = "";
    // One entry per character of `text`; null at a "\n" paragraph mark, a
    // "\t" tab, or a "\v" line break — characters with no w:t to select or
    // replace. A match containing one keeps it through a replacement, the
    // paragraph-mark rule extended.
    const refs: (CharRef | null)[] = [];
    for (const para of unit.paragraphs) {
      if (refs.length > 0) {
        text += "\n";
        refs.push(null);
      }
      for (const run of runsOf(para)) {
        for (const c of run.content) {
          if (c.kind === "tab") {
            text += "\t";
            refs.push(null);
          } else if (c.kind === "break" && c.breakType === "line") {
            text += "\v";
            refs.push(null);
          } else if (c.kind === "text" && c.srcT) {
            const t = c.srcT as XmlElement;
            for (let i = 0; i < c.text.length; i++) {
              text += c.text[i];
              refs.push({ t, offset: i });
            }
          }
        }
      }
    }

    // Match spans over `text`, non-overlapping, left to right in both paths.
    const spans: [number, number][] = [];
    if ("literal" in compiled) {
      const hay = opts?.matchCase ? text : text.toLowerCase();
      const q = compiled.literal;
      let from = 0;
      for (;;) {
        const idx = hay.indexOf(q, from);
        if (idx === -1) break;
        from = idx + q.length;
        if (opts?.wholeWord && (isWordChar(text[idx - 1]) || isWordChar(text[idx + q.length]))) continue;
        spans.push([idx, idx + q.length]);
      }
    } else {
      compiled.re.lastIndex = 0;
      for (;;) {
        const m = compiled.re.exec(text);
        if (!m) break;
        // A zero-length match selects nothing; step past it rather than loop.
        if (m[0].length === 0) {
          compiled.re.lastIndex++;
          continue;
        }
        spans.push([m.index, m.index + m[0].length]);
      }
    }

    for (const [from, to] of spans) {
      // Convert char span -> per-t ranges (paragraph marks, tabs and line
      // breaks have none).
      const ranges: FormattedRange[] = [];
      for (let i = from; i < to; i++) {
        const ref = refs[i];
        if (!ref) continue;
        const last = ranges[ranges.length - 1];
        if (last && last.t === ref.t && last.end === ref.offset) last.end = ref.offset + 1;
        else ranges.push({ t: ref.t, start: ref.offset, end: ref.offset + 1 });
      }
      // A match must cover at least one real text character: a hit made only
      // of paragraph marks/tabs has nothing to select or replace.
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
