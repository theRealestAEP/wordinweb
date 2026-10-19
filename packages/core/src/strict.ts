/**
 * Strict Open XML (ISO/IEC 29500-1 Strict, Word's "Strict Open XML
 * Document") read as Transitional. A Strict package is the same package
 * with other namespaces and relationship types (purl.oclc.org/ooxml/...),
 * measures that carry a unit ("72pt", "2.54cm") where Transitional writes
 * an integer in the attribute's own unit, percentages as "63%", and a few
 * renamed values and elements. The parsers, the layout, the compare and
 * the writers all speak Transitional, so a Strict part is rewritten to
 * Transitional once, as its bytes are read, and nothing downstream knows
 * the difference; whatever is saved or compared is written as Transitional.
 *
 * The rewrite walks the part's tags and touches attribute values only:
 * text content is never read. Every reader of a package goes through here:
 * DocxDocument.load (a package in memory), the streaming reader
 * (stream/body.ts, the PDF renderer) and compare-cli's input reader.
 */

/** Every Strict namespace and relationship type starts with this. */
const STRICT_URI = "http://purl.oclc.org/ooxml/";
const TRANSITIONAL_URI = "http://schemas.openxmlformats.org/";

/** Names that differ beyond the "2006" the Transitional URIs carry. */
const URI_RENAMES: Record<string, string> = {
  extendedProperties: "extended-properties",
  customProperties: "custom-properties",
};

/**
 * A Transitional URI for a Strict one: the family keeps its name and gains
 * the 2006 edition (http://purl.oclc.org/ooxml/wordprocessingml/main ->
 * http://schemas.openxmlformats.org/wordprocessingml/2006/main); a
 * relationship type keeps its kind.
 */
export function transitionalUri(uri: string): string {
  if (!uri.startsWith(STRICT_URI)) return uri;
  const path = uri.slice(STRICT_URI.length);
  const slash = path.indexOf("/");
  if (slash === -1) return uri;
  const family = path.slice(0, slash);
  const rest = path.slice(slash + 1);
  const kind = rest.startsWith("relationships/") ? rest.slice("relationships/".length) : rest;
  const renamed = URI_RENAMES[kind] ?? kind;
  return `${TRANSITIONAL_URI}${family}/2006/${rest.startsWith("relationships/") ? `relationships/${renamed}` : renamed}`;
}

const UNIVERSAL_MEASURE = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/;
const PERCENTAGE = /^(-?\d+(?:\.\d+)?)%$/;
const POINTS_PER: Record<string, number> = { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 };

/** Elements whose w:val is a half-point measure (ST_HpsMeasure, ST_SignedHpsMeasure). */
const HALF_POINT_VAL = new Set(["sz", "szCs", "position", "kern", "hps", "hpsRaise", "hpsBaseText"]);

/** Elements whose w:w is a width or a percentage (ST_MeasurementOrPercent), which Word writes in fiftieths of a percent. */
const WIDTH_OR_PERCENT = new Set(["tblW", "tcW", "tblInd", "tblCellSpacing", "wBefore", "wAfter", "top", "bottom", "left", "right", "start", "end"]);

/** Elements whose w:start and w:end children Transitional names w:left and w:right. */
const SIDED = new Set(["pBdr", "pgBorders", "divBdr", "tcBorders", "tblBorders", "tblCellMar", "tcMar"]);

const JUSTIFICATION: Record<string, string> = { start: "left", end: "right" };
const TEXT_DIRECTION: Record<string, string> = { tb: "lrTb", rl: "tbRl", lr: "btLr", tbV: "lrTbV", rlV: "tbRlV", lrV: "tbLrV" };

/** Prefixes of the DrawingML family, whose measures are EMUs and whose percentages are thousandths. */
const DRAWING_PREFIXES = new Set(["a", "wp", "pic", "wpg", "wps", "wpc", "dgm", "cdr", "lc", "xdr", "dsp"]);

/**
 * A Strict attribute value as Transitional writes it, by the element and
 * attribute it sits on; null drops the attribute. Namespaces are told by
 * their customary prefixes, as the parsers tell them.
 */
function transitionalValue(prefix: string, element: string, name: string, value: string): string | null {
  if (name === "conformance" && element === "document") return null;
  if (prefix === "w") {
    if ((element === "jc" || element === "lvlJc" || element === "tab") && name === "val") return JUSTIFICATION[value] ?? value;
    if (element === "textDirection" && name === "val") return TEXT_DIRECTION[value] ?? value;
  }
  const percent = PERCENTAGE.exec(value);
  if (percent) {
    const n = Number(percent[1]);
    if (prefix === "w") return String(Math.round(name === "w" && WIDTH_OR_PERCENT.has(element) ? n * 50 : n));
    if (DRAWING_PREFIXES.has(prefix)) return String(Math.round(n * 1000));
    if (prefix === "c") return String(Math.round(n));
    return value;
  }
  const measure = UNIVERSAL_MEASURE.exec(value);
  if (!measure) return value;
  const points = Number(measure[1]) * POINTS_PER[measure[2]];
  if (prefix === "w" || prefix === "m") {
    if (name === "val" && HALF_POINT_VAL.has(element)) return String(Math.round(points * 2));
    // A border's width is in eighths of a point, its padding in points.
    if (name === "sz") return String(Math.round(points * 8));
    if (name === "space" && element !== "cols" && element !== "col") return String(Math.round(points));
    return String(Math.round(points * 20));
  }
  if (DRAWING_PREFIXES.has(prefix)) return String(Math.round(points * 12700));
  return value;
}

/** An open element: its name as read, and as written. */
interface Open {
  name: string;
  out: string;
}

/** The rewrite of one part, across the pieces it comes in. */
export interface RewriteState {
  /** Whether the part is Strict: undefined until its root tag is seen. */
  strict: boolean | undefined;
  stack: Open[];
}

export function newRewriteState(): RewriteState {
  return { strict: undefined, stack: [] };
}

const ATTRIBUTE = /(\s+)([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function split(qualified: string): [string, string] {
  const colon = qualified.indexOf(":");
  return colon === -1 ? ["", qualified] : [qualified.slice(0, colon), qualified.slice(colon + 1)];
}

/** A start tag rewritten, or itself when nothing in it changes. */
function rewriteStartTag(tag: string, state: RewriteState): { tag: string; name: string; out: string; empty: boolean } {
  const empty = tag.endsWith("/>");
  const nameEnd = tag.search(/[\s/>]/);
  const name = tag.slice(1, nameEnd);
  const [prefix, local] = split(name);
  if (state.strict === undefined) state.strict = tag.includes(STRICT_URI);
  let out = name;
  const parent = state.stack[state.stack.length - 1];
  if (state.strict && prefix === "w" && parent && SIDED.has(split(parent.name)[1])) {
    if (local === "start") out = "w:left";
    else if (local === "end") out = "w:right";
  }
  let changed = out !== name;
  const attributes = tag.slice(nameEnd, empty ? -2 : -1).replace(ATTRIBUTE, (whole, space: string, attribute: string, dq: string | undefined, sq: string | undefined) => {
    const value = dq ?? sq ?? "";
    const [, attrLocal] = split(attribute);
    let next: string | null = value;
    // A namespace declaration, a relationship's Type, a graphicData's or schemaRef's uri.
    if (attribute === "Type" || attribute.startsWith("xmlns") || attrLocal === "uri") next = transitionalUri(value);
    else if (state.strict) next = transitionalValue(prefix, local, attrLocal, value);
    if (next === value) return whole;
    changed = true;
    return next === null ? "" : `${space}${attribute}="${next}"`;
  });
  return { tag: changed ? `<${out}${attributes}${empty ? "/>" : ">"}` : tag, name, out, empty };
}

/**
 * Rewrite the complete tags in `text`; `rest` is the start of a tag the text
 * ends inside, to go before the next piece.
 */
export function rewriteTags(text: string, state: RewriteState): { out: string; rest: string } {
  const pieces: string[] = [];
  let i = 0;
  const n = text.length;
  while (i < n) {
    const lt = text.indexOf("<", i);
    if (lt === -1) break;
    pieces.push(text.slice(i, lt));
    const c = text[lt + 1];
    let end: number;
    if (c === "?") {
      end = text.indexOf("?>", lt + 2);
      if (end === -1) return { out: pieces.join(""), rest: text.slice(lt) };
      end += 2;
      pieces.push(text.slice(lt, end));
    } else if (c === "!") {
      const close = text.startsWith("<!--", lt) ? "-->" : text.startsWith("<![CDATA[", lt) ? "]]>" : ">";
      end = text.indexOf(close, lt + 2);
      if (end === -1) return { out: pieces.join(""), rest: text.slice(lt) };
      end += close.length;
      pieces.push(text.slice(lt, end));
    } else if (c === "/") {
      end = text.indexOf(">", lt + 2);
      if (end === -1) return { out: pieces.join(""), rest: text.slice(lt) };
      end += 1;
      const open = state.stack.pop();
      pieces.push(open && open.out !== open.name ? `</${open.out}>` : text.slice(lt, end));
    } else {
      end = tagEnd(text, lt);
      if (end === -1) return { out: pieces.join(""), rest: text.slice(lt) };
      const start = rewriteStartTag(text.slice(lt, end), state);
      pieces.push(start.tag);
      if (!start.empty) state.stack.push({ name: start.name, out: start.out });
    }
    i = end;
  }
  pieces.push(text.slice(i));
  return { out: pieces.join(""), rest: "" };
}

/** A part's XML as Transitional: itself when it holds nothing Strict. */
export function normalizeStrictXml(text: string): string {
  if (!text.includes(STRICT_URI)) return text;
  const { out, rest } = rewriteTags(text, newRewriteState());
  return out + rest;
}

const EMPTY = new Uint8Array(0);

/**
 * A part's bytes rewritten to Transitional as they stream by, in pieces cut
 * anywhere. The root tag decides: a part that declares no Strict namespace
 * passes through untouched, and undecoded.
 */
export class StrictRewriter {
  private readonly state = newRewriteState();
  private readonly decoder = new TextDecoder();
  private readonly encoder = new TextEncoder();
  /** Raw pieces held until the root tag decides. */
  private held: Uint8Array[] = [];
  /** Decoded text not yet rewritten. */
  private text = "";

  /** Whether the part is Strict: undefined until its root tag has come. */
  get strict(): boolean | undefined {
    return this.state.strict;
  }

  push(piece: Uint8Array): Uint8Array {
    if (this.state.strict === false) return piece;
    this.text += this.decoder.decode(piece, { stream: true });
    if (this.state.strict === undefined) {
      this.held.push(piece);
      const root = rootTag(this.text);
      if (root === null) return EMPTY;
      this.state.strict = root.includes(STRICT_URI);
      const held = this.held;
      this.held = [];
      if (!this.state.strict) {
        this.text = "";
        return concat(held);
      }
    }
    return this.rewrite();
  }

  /** The last of the part, once every piece is in. */
  end(): Uint8Array {
    if (this.state.strict === false) return EMPTY;
    this.text += this.decoder.decode();
    if (this.state.strict === undefined) {
      // No root tag at all: nothing to tell Strict by.
      const held = this.held;
      this.held = [];
      this.text = "";
      return concat(held);
    }
    const out = this.rewrite();
    const rest = this.encoder.encode(this.text);
    this.text = "";
    return concat([out, rest]);
  }

  private rewrite(): Uint8Array {
    const { out, rest } = rewriteTags(this.text, this.state);
    this.text = rest;
    return this.encoder.encode(out);
  }

  /** `pull` (a piece per call, null at the end) with every piece rewritten. */
  through(pull: () => Uint8Array | null): () => Uint8Array | null {
    let done = false;
    return () => {
      while (!done) {
        const piece = pull();
        if (piece === null) {
          done = true;
          const last = this.end();
          return last.length > 0 ? last : null;
        }
        const out = this.push(piece);
        if (out.length > 0) return out;
      }
      return null;
    };
  }
}

/** Just past a start tag's >, the first outside a quoted value; -1 while the text ends inside the tag. */
function tagEnd(text: string, lt: number): number {
  for (let j = lt + 1; j < text.length; j++) {
    const ch = text[j];
    if (ch === '"' || ch === "'") {
      j = text.indexOf(ch, j + 1);
      if (j === -1) return -1;
    } else if (ch === ">") return j + 1;
  }
  return -1;
}

/** The first element's start tag, or null until the text holds it whole. */
function rootTag(text: string): string | null {
  let at = 0;
  for (;;) {
    const lt = text.indexOf("<", at);
    if (lt === -1) return null;
    const c = text[lt + 1];
    if (c !== "?" && c !== "!") {
      const end = tagEnd(text, lt);
      return end === -1 ? null : text.slice(lt, end);
    }
    const close = c === "?" ? "?>" : text.startsWith("<!--", lt) ? "-->" : ">";
    const end = text.indexOf(close, lt + 2);
    if (end === -1) return null;
    at = end + close.length;
  }
}

/**
 * The text parts of a package in memory rewritten in place to Transitional,
 * when its root relationships or its main document part are Strict. Returns
 * whether they were.
 */
export function normalizeStrictPackage(files: Record<string, Uint8Array>): boolean {
  const decoder = new TextDecoder();
  const head = (name: string): string => (files[name] ? decoder.decode(files[name].subarray(0, 4096)) : "");
  if (!head("_rels/.rels").includes(STRICT_URI) && !head("word/document.xml").includes(STRICT_URI)) return false;
  const encoder = new TextEncoder();
  for (const name in files) {
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const text = decoder.decode(files[name]);
    const out = normalizeStrictXml(text);
    if (out !== text) files[name] = encoder.encode(out);
  }
  return true;
}

function concat(pieces: Uint8Array[]): Uint8Array {
  if (pieces.length === 1) return pieces[0];
  const out = new Uint8Array(pieces.reduce((total, piece) => total + piece.length, 0));
  let at = 0;
  for (const piece of pieces) {
    out.set(piece, at);
    at += piece.length;
  }
  return out;
}
