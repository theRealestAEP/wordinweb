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
  /** Word tokens: the `w:t` and the slice inside it. */
  t?: XmlElement;
  start?: number;
  end?: number;
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
   * True when every child of the paragraph is `w:pPr`, `w:r` or a passenger.
   * A paragraph holding a hyperlink, a content control, a smart tag or inline
   * math is NOT word-diffable here: its content would have to be rebuilt, and
   * a revision wrapper may not sit around a `w:hyperlink` (the schema puts
   * `w:ins` inside it, not around it). Such a paragraph is compared whole.
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

  for (const cEl of pEl.children) {
    const ln = localName(cEl.name);
    if (ln === "pPr") continue;
    if (PASSENGERS.has(ln)) {
      pending.push(cEl);
      continue;
    }
    if (ln !== "r") {
      inlineDiffable = false;
      continue;
    }
    for (const rc of cEl.children) {
      const rln = localName(rc.name);
      if (rln === "rPr") continue;
      if (rln === "t") {
        let at = 0;
        for (const word of wordTokens(rc.text)) {
          push({ key: "w\u0000" + keyOf(word), text: word, run: cEl, t: rc, start: at, end: at + word.length });
          at += word.length;
        }
        // An empty w:t contributes nothing to the alignment, which is right:
        // it is a placeholder, not content.
        continue;
      }
      push({ key: "a\u0000" + serializeXml(rc), text: "", run: cEl, atom: rc });
    }
  }

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
export function sliceContent(
  tokens: readonly ContentToken[],
  from: number,
  to: number,
  deleted: boolean,
): XmlElement[] {
  const out: XmlElement[] = [];
  let i = from;
  while (i < to) {
    const token = tokens[i];
    if (!deleted && token.lead) out.push(...token.lead.map(cloneXml));
    const run = token.run;
    const rPr = run.children.find((c) => localName(c.name) === "rPr");
    const children: XmlElement[] = rPr ? [cloneXml(rPr)] : [];
    // One output run per source run, so run properties are never lost and
    // never invented.
    while (i < to && tokens[i].run === run) {
      const cur = tokens[i];
      if (cur.atom) {
        children.push(cloneXml(cur.atom));
        i++;
        continue;
      }
      // Consecutive word tokens out of the same w:t are contiguous by
      // construction, so they collapse back into one text element.
      const t = cur.t!;
      let text = "";
      while (i < to && tokens[i].run === run && tokens[i].t === t) {
        text += tokens[i].text;
        i++;
      }
      children.push(textElement(t, text, deleted));
    }
    out.push({ name: run.name, attrs: { ...run.attrs }, children, text: "" });
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
