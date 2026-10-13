import { XmlElement, cloneXml, localName, serializeXml } from "../../xml.js";
import { wordTokens } from "./text.js";

/**
 * Paragraph content as a flat token sequence the inner diff can align, plus
 * the machinery to rebuild a slice of it as fresh XML.
 *
 * A token is either a WORD (a run of non-space plus its trailing whitespace,
 * addressed as a slice of one `w:t`) or an ATOM — any run child that is not
 * text: a break, a tab, a drawing, a field character, a footnote reference. An
 * atom matches only itself, byte for byte, so a changed image is a change and
 * an unchanged one never is.
 */

/** Elements that carry no visible content and must never read as a difference. */
const PASSENGERS = new Set(["proofErr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd"]);

/** True for an element that rides along with the content instead of being diffed. */
export function isPassenger(ln: string): boolean {
  return PASSENGERS.has(ln);
}

export interface ContentToken {
  /** What the LCS compares. Words key on their own text; atoms on their XML. */
  key: string;
  /** The text this token contributes (empty for an atom). */
  text: string;
  /** The `w:r` this token came from. */
  run: XmlElement;
  /**
   * The run containers the token sits inside, outermost first: hyperlinks,
   * content controls, simple fields, smart tags. Revisions go INSIDE them
   * (the schema puts w:ins inside a w:hyperlink, not around it), so the
   * emitter rebuilds each container around its own words.
   */
  path: readonly XmlElement[];
  /** Word tokens: the `w:t` and the slice inside it (the first piece's). */
  t?: XmlElement;
  start?: number;
  end?: number;
  /**
   * Word tokens: every `w:t` slice the word covers, in order. A word Word
   * split across runs that share the same properties ("Purp" + "ose") is
   * one token with two pieces, so it matches "Purpose" on the other side.
   */
  pieces?: { run: XmlElement; t: XmlElement; start: number; end: number; text: string }[];
  /** Atom tokens: the element itself. */
  atom?: XmlElement;
  /**
   * Passenger elements that sat immediately before this token as direct
   * children of the paragraph. They ride along with whichever slice emits the
   * token and never count as a difference of their own.
   */
  lead?: XmlElement[];
}

export interface ParagraphContent {
  tokens: ContentToken[];
  /** Passengers after the last token. */
  tail: XmlElement[];
  /** The paragraph's full text. */
  text: string;
  /**
   * False when the paragraph holds content no word diff can rebuild, inline
   * math above all; such a paragraph is compared whole. Runs inside run
   * containers (CONTAINERS) are diffable: each token records its path.
   */
  inlineDiffable: boolean;
}

export interface ReadOptions {
  /**
   * Whether a whitespace-only difference counts. With it off, a token keys on
   * its word alone, so "double space becomes single" leaves the paragraph
   * matching — Word offers the same switch. It is lossy on purpose: a
   * difference that is not recorded cannot be rejected back either.
   */
  whitespace?: boolean;
}

const WORD_CHAR = /[\p{L}\p{M}\p{N}]/u;

/** Run containers a word diff descends into; revisions sit inside them. */
export const CONTAINERS = new Set(["hyperlink", "smartTag", "fldSimple", "customXml", "dir", "bdo", "sdt"]);
/** A container's own property elements: kept with the container, not diffed. */
const CONTAINER_PROPS = new Set(["pPr", "sdtPr", "sdtEndPr", "smartTagPr", "customXmlPr"]);

/** Read a `w:p` into the token sequence the inner diff aligns. */
export function readParagraph(pEl: XmlElement, options: ReadOptions = {}): ParagraphContent {
  const keyOf = options.whitespace === false ? (word: string): string => word.trimEnd() : (word: string): string => word;
  const tokens: ContentToken[] = [];
  let pending: XmlElement[] = [];
  let inlineDiffable = true;

  const push = (token: ContentToken): void => {
    if (pending.length > 0) {
      token.lead = pending;
      pending = [];
    }
    tokens.push(token);
  };

  // Text is tokenized per SEGMENT: the w:t elements of consecutive runs with
  // nothing but text between them, whatever their properties. Word splits
  // runs for reasons no reader sees (editing sessions, proofing marks), and
  // tokenized run by run a word cut in two ("Purp" + "ose") never matched
  // "Purpose" and read as a change. Word boundaries must not depend on run
  // boundaries at all: formatting is compared piece by piece instead
  // (emitEqualSpan), so a bold "Purpose:" against bold "Purpose" + plain ":"
  // records only the colon's lost bold.
  type Piece = { run: XmlElement; t: XmlElement; text: string };
  let segment: Piece[] = [];
  let segmentPath: readonly XmlElement[] = [];
  const flush = (): void => {
    if (segment.length === 0) return;
    const text = segment.map((p) => p.text).join("");
    // Letters and digits run on across a run boundary only when both runs
    // carry the same properties: a word glued to a superscript note number
    // ("הלשון" + "1") is two tokens, so editing the number does not mark the
    // word. Punctuation and spaces are tokens of their own either way.
    const cuts: number[] = [];
    let base = 0;
    for (let i = 0; i + 1 < segment.length; i++) {
      base += segment[i].text.length;
      if (WORD_CHAR.test(text[base - 1] ?? "") && WORD_CHAR.test(text[base] ?? "") && runPropsKey(segment[i].run) !== runPropsKey(segment[i + 1].run)) {
        cuts.push(base);
      }
    }
    const words: string[] = [];
    let from = 0;
    for (const cut of [...cuts, text.length]) {
      words.push(...wordTokens(text.slice(from, cut)));
      from = cut;
    }
    let at = 0;
    for (const word of words) {
      const end = at + word.length;
      const pieces: NonNullable<ContentToken["pieces"]> = [];
      let base = 0;
      for (const p of segment) {
        const from = Math.max(at, base);
        const to = Math.min(end, base + p.text.length);
        if (from < to) pieces.push({ run: p.run, t: p.t, start: from - base, end: to - base, text: text.slice(from, to) });
        base += p.text.length;
      }
      const first = pieces[0];
      push({ key: "w\u0000" + keyOf(word), text: word, run: first.run, t: first.t, start: first.start, end: first.end, pieces, path: segmentPath });
      at = end;
    }
    // An empty w:t contributes nothing to the alignment, which is right:
    // it is a placeholder, not content.
    segment = [];
  };

  const walk = (kids: readonly XmlElement[], path: readonly XmlElement[]): void => {
    for (const cEl of kids) {
      const ln = localName(cEl.name);
      if (CONTAINER_PROPS.has(ln)) continue;
      if (PASSENGERS.has(ln)) {
        flush();
        pending.push(cEl);
        continue;
      }
      if (CONTAINERS.has(ln)) {
        // A content control's words sit in its w:sdtContent; the others hold
        // their runs directly.
        const inner = ln === "sdt" ? cEl.children.find((c) => localName(c.name) === "sdtContent") : cEl;
        if (!inner) continue;
        flush();
        walk(inner.children, [...path, cEl]);
        flush();
        continue;
      }
      if (ln !== "r") {
        flush();
        inlineDiffable = false;
        continue;
      }
      if (segment.length > 0 && segmentPath !== path) flush();
      segmentPath = path;
      for (const rc of cEl.children) {
        const rln = localName(rc.name);
        if (rln === "rPr") continue;
        if (rln === "t") {
          segment.push({ run: cEl, t: rc, text: rc.text });
          continue;
        }
        flush();
        push({ key: "a\u0000" + serializeXml(rc), text: "", run: cEl, atom: rc, path });
      }
    }
  };
  walk(pEl.children, []);
  flush();

  return { tokens, tail: pending, text: paragraphText(pEl), inlineDiffable };
}

/** Every `w:t` under an element, concatenated in document order. */
export function paragraphText(el: XmlElement): string {
  let out = "";
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "t") {
      out += e.text;
      return;
    }
    for (const c of e.children) walk(c);
  };
  for (const c of el.children) {
    if (localName(c.name) !== "pPr") walk(c);
  }
  return out;
}

/**
 * Rebuild `tokens[from, to)` as fresh XML: cloned runs carrying exactly the
 * text those tokens cover.
 *
 * `deleted` retypes every `w:t` as `w:delText`, which is what a `w:del` must
 * hold, and drops passengers — a bookmark inside struck text has nothing left
 * to point at.
 */
/** An emitted element and the containers it belongs inside. */
export interface Placed {
  el: XmlElement;
  path: readonly XmlElement[];
}

export function sliceItems(
  tokens: readonly ContentToken[],
  from: number,
  to: number,
  deleted: boolean,
): Placed[] {
  const out: Placed[] = [];
  // One output run per source run, so run properties are never lost and
  // never invented; consecutive slices of one w:t collapse back into one
  // text element. A token spanning runs contributes a piece to each.
  let open: { run: XmlElement; el: XmlElement } | null = null;
  let text: { t: XmlElement; el: XmlElement } | null = null;
  const runFor = (run: XmlElement, path: readonly XmlElement[]): XmlElement => {
    if (open?.run === run) return open.el;
    const rPr = run.children.find((c) => localName(c.name) === "rPr");
    const el: XmlElement = { name: run.name, attrs: { ...run.attrs }, children: rPr ? [cloneXml(rPr)] : [], text: "" };
    out.push({ el, path });
    open = { run, el };
    text = null;
    return el;
  };
  for (let i = from; i < to; i++) {
    const token = tokens[i];
    if (!deleted && token.lead) {
      out.push(...token.lead.map((el) => ({ el: cloneXml(el), path: token.path })));
      open = null;
      text = null;
    }
    if (token.atom) {
      runFor(token.run, token.path).children.push(cloneXml(token.atom));
      text = null;
      continue;
    }
    for (const piece of token.pieces ?? []) {
      const el = runFor(piece.run, token.path);
      if (text && text.t === piece.t) {
        text.el.text += piece.text;
        continue;
      }
      const te = textElement(piece.t, piece.text, deleted);
      el.children.push(te);
      text = { t: piece.t, el: te };
    }
  }
  return out;
}

function textElement(t: XmlElement, text: string, deleted: boolean): XmlElement {
  const name = deleted ? t.name.replace(/t$/, "delText") : t.name;
  return { name, attrs: { ...t.attrs, "xml:space": "preserve" }, children: [], text };
}

/** The `w:rPr` of a run, serialized, as an equality key ("" when absent). */
export function runPropsKey(run: XmlElement): string {
  const rPr = run.children.find((c) => localName(c.name) === "rPr");
  return rPr ? serializeXml(rPr) : "";
}
