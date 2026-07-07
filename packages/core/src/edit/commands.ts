import { DocxDocument } from "../docx.js";
import { Run, RunProps } from "../model.js";
import { XmlElement, cloneXml, localName, child } from "../xml.js";

/**
 * Editing commands, v1: character formatting over a selection.
 *
 * The XML tree is the source of truth. Commands split `w:r` elements at
 * selection boundaries, mutate `w:rPr`, then `doc.refresh()` re-derives the
 * model. Everything untouched round-trips byte-for-byte on save.
 */

/** A contiguous character range of a single run covered by the selection. */
export interface SelectionSegment {
  run: Run;
  /** Target w:t element; null → format the whole run (fields, images). */
  t: XmlElement | null;
  /** Char offsets within t's text (ignored when t is null). */
  start: number;
  end: number;
  /** Effective props of the covered text (for toggle decisions in the UI). */
  props: RunProps;
}

export interface RunFormatPatch {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** "#RRGGBB"; null removes the direct color. */
  color?: string | null;
  /** Word highlight name ("yellow", "cyan", …); null removes. */
  highlight?: string | null;
  /** Font size in points. */
  fontSizePt?: number;
  fontFamily?: string;
  /** Superscript/subscript; null returns the run to the baseline. */
  verticalAlign?: "superscript" | "subscript" | null;
  /** Remove all direct character formatting (and character style). */
  clear?: boolean;
}

/** A formatted char-range in the post-edit XML (for re-selection). */
export interface FormattedRange {
  t: XmlElement;
  start: number;
  end: number;
}

export function applyRunFormat(
  doc: DocxDocument,
  segments: SelectionSegment[],
  patch: RunFormatPatch,
): FormattedRange[] {
  const formatted: FormattedRange[] = [];
  // Group by run; merge ranges on the same w:t.
  const byRun = new Map<Run, SelectionSegment[]>();
  for (const seg of segments) {
    if (!seg.run.src) continue;
    const list = byRun.get(seg.run) ?? [];
    list.push(seg);
    byRun.set(seg.run, list);
  }

  for (const [run, segs] of byRun) {
    const rEl = run.src!;
    const parent = run.srcParent;
    const tTargets = new Map<XmlElement, { start: number; end: number }>();
    let wholeRun = false;
    for (const seg of segs) {
      if (!seg.t) {
        wholeRun = true;
        continue;
      }
      const cur = tTargets.get(seg.t);
      if (cur) {
        cur.start = Math.min(cur.start, seg.start);
        cur.end = Math.max(cur.end, seg.end);
      } else {
        tTargets.set(seg.t, { start: seg.start, end: seg.end });
      }
    }

    // Full coverage of every text child → format the run in place.
    const fullyCovered =
      wholeRun ||
      Array.from(tTargets).every(([t, r]) => r.start <= 0 && r.end >= t.text.length);
    const coversAllTs =
      wholeRun ||
      (tTargets.size >= countTextChildren(rEl) && fullyCovered);

    if (coversAllTs || !parent || tTargets.size !== 1) {
      // Whole-run formatting (also the safe fallback for multi-t partials).
      setRunProps(rEl, patch);
      for (const c of rEl.children) {
        if (localName(c.name) === "t") formatted.push({ t: c, start: 0, end: c.text.length });
      }
      continue;
    }

    const [t, range] = Array.from(tTargets)[0];
    const middleT = splitAndFormat(parent, rEl, t, range.start, range.end, patch);
    if (middleT) formatted.push({ t: middleT, start: 0, end: middleT.text.length });
  }

  doc.refresh();
  return formatted;
}

function countTextChildren(rEl: XmlElement): number {
  return rEl.children.filter((c) => localName(c.name) === "t").length;
}

/**
 * Split `rEl` (inside parent.children) around [start,end) of text child `t`
 * into up to three runs; apply the patch to the middle one.
 */
function splitAndFormat(
  parent: XmlElement,
  rEl: XmlElement,
  t: XmlElement,
  start: number,
  end: number,
  patch: RunFormatPatch,
): XmlElement | null {
  const idx = parent.children.indexOf(rEl);
  const tIdx = rEl.children.indexOf(t);
  if (idx === -1 || tIdx === -1) return null;

  const text = t.text;
  start = Math.max(0, Math.min(start, text.length));
  end = Math.max(start, Math.min(end, text.length));
  if (start === end) return null;

  const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  const prefix = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";

  const makeRun = (children: XmlElement[]): XmlElement => ({
    name: rEl.name,
    attrs: { ...rEl.attrs },
    text: "",
    children: [...(rPr ? [cloneXml(rPr)] : []), ...children],
  });
  const makeT = (s: string): XmlElement => ({
    name: t.name,
    attrs: { ...t.attrs, "xml:space": "preserve" },
    text: s,
    children: [],
  });

  const before = rEl.children.slice(0, tIdx).filter((c) => localName(c.name) !== "rPr");
  const after = rEl.children.slice(tIdx + 1);

  const newRuns: XmlElement[] = [];
  const beforeChildren = [...before];
  if (start > 0) beforeChildren.push(makeT(text.slice(0, start)));
  if (beforeChildren.length > 0) newRuns.push(makeRun(beforeChildren));

  const middleT = makeT(text.slice(start, end));
  const middle = makeRun([middleT]);
  setRunProps(middle, patch);
  newRuns.push(middle);

  const afterChildren: XmlElement[] = [];
  if (end < text.length) afterChildren.push(makeT(text.slice(end)));
  afterChildren.push(...after);
  if (afterChildren.length > 0) newRuns.push(makeRun(afterChildren));

  parent.children.splice(idx, 1, ...newRuns);
  void prefix;
  return middleT;
}

// ---------- rPr mutation ----------

/** Schema-ish ordering for rPr children so Word accepts the result. */
const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
  "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
  "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
  "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
  "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
];

function prefixOf(el: XmlElement): string {
  return el.name.includes(":") ? el.name.slice(0, el.name.indexOf(":") + 1) : "";
}

function ensureRPr(rEl: XmlElement): XmlElement {
  let rPr = rEl.children.find((c) => localName(c.name) === "rPr");
  if (!rPr) {
    rPr = { name: prefixOf(rEl) + "rPr", attrs: {}, children: [], text: "" };
    rEl.children.unshift(rPr);
  }
  return rPr;
}

/** Set (or remove, when attrs === null) a property child of rPr. */
function setProp(
  rPr: XmlElement,
  local: string,
  attrs: Record<string, string> | null,
): void {
  const idx = rPr.children.findIndex((c) => localName(c.name) === local);
  if (attrs === null) {
    if (idx !== -1) rPr.children.splice(idx, 1);
    return;
  }
  const prefix = prefixOf(rPr);
  const el: XmlElement = { name: prefix + local, attrs: prefixAttrs(prefix, attrs), children: [], text: "" };
  if (idx !== -1) {
    rPr.children[idx] = el;
    return;
  }
  const orderIdx = RPR_ORDER.indexOf(local);
  let insertAt = rPr.children.length;
  for (let i = 0; i < rPr.children.length; i++) {
    const iOrder = RPR_ORDER.indexOf(localName(rPr.children[i].name));
    if (iOrder !== -1 && orderIdx !== -1 && iOrder > orderIdx) {
      insertAt = i;
      break;
    }
  }
  rPr.children.splice(insertAt, 0, el);
}

function prefixAttrs(prefix: string, attrs: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    out[k.includes(":") ? k : prefix + k] = v;
  }
  return out;
}

export function setRunProps(rEl: XmlElement, patch: RunFormatPatch): void {
  if (patch.clear) {
    rEl.children = rEl.children.filter((c) => localName(c.name) !== "rPr");
    return;
  }
  const rPr = ensureRPr(rEl);
  if (patch.bold !== undefined) {
    setProp(rPr, "b", patch.bold ? {} : { val: "0" });
    setProp(rPr, "bCs", patch.bold ? {} : { val: "0" });
  }
  if (patch.italic !== undefined) {
    setProp(rPr, "i", patch.italic ? {} : { val: "0" });
    setProp(rPr, "iCs", patch.italic ? {} : { val: "0" });
  }
  if (patch.underline !== undefined) {
    setProp(rPr, "u", { val: patch.underline ? "single" : "none" });
  }
  if (patch.strike !== undefined) {
    setProp(rPr, "strike", patch.strike ? {} : { val: "0" });
  }
  if (patch.color !== undefined) {
    setProp(rPr, "color", patch.color === null ? null : { val: patch.color.replace(/^#/, "").toUpperCase() });
  }
  if (patch.highlight !== undefined) {
    setProp(rPr, "highlight", patch.highlight === null ? null : { val: patch.highlight });
  }
  if (patch.fontSizePt !== undefined) {
    const half = String(Math.round(patch.fontSizePt * 2));
    setProp(rPr, "sz", { val: half });
    setProp(rPr, "szCs", { val: half });
  }
  if (patch.fontFamily !== undefined) {
    setProp(rPr, "rFonts", { ascii: patch.fontFamily, hAnsi: patch.fontFamily, cs: patch.fontFamily });
  }
  if (patch.verticalAlign !== undefined) {
    setProp(rPr, "vertAlign", patch.verticalAlign === null ? null : { val: patch.verticalAlign });
  }
}

// ---------- selection format summary ----------

export interface SelectionFormat {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** Common font size in points, if uniform. */
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  fontFamily?: string;
  verticalAlign?: "superscript" | "subscript";
}

/** Summarize effective formatting across segments (for toolbar state/toggles). */
export function summarizeSelection(segments: SelectionSegment[]): SelectionFormat | null {
  if (segments.length === 0) return null;
  const all = (f: (p: RunProps) => boolean | undefined) => segments.every((s) => f(s.props) === true);
  const first = segments[0].props;
  const uniform = <T>(f: (p: RunProps) => T | undefined): T | undefined => {
    const v = f(first);
    return segments.every((s) => f(s.props) === v) ? v : undefined;
  };
  return {
    bold: all((p) => p.bold),
    italic: all((p) => p.italic),
    underline: all((p) => (p.underline !== undefined && p.underline !== "none") || undefined),
    strike: all((p) => p.strike),
    fontSizePt: mapDefined(uniform((p) => p.size), (px) => Math.round(((px * 3) / 4) * 2) / 2),
    color: uniform((p) => p.color),
    highlight: uniform((p) => p.highlight),
    verticalAlign: mapDefined(
      uniform((p) => p.verticalAlign),
      (v) => (v === "superscript" || v === "subscript" ? v : undefined),
    ),
    fontFamily: uniform((p) => p.font),
  };
}

function mapDefined<T, U>(v: T | undefined, f: (v: T) => U): U | undefined {
  return v === undefined ? undefined : f(v);
}
