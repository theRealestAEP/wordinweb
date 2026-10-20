/**
 * Minimal, fast XML parser producing a lightweight element tree.
 *
 * OOXML parts are machine-generated, namespace-prefixed, well-formed XML with
 * no DTDs. This parser handles exactly that subset: elements, attributes,
 * text, CDATA, comments, processing instructions, and the built-in + numeric
 * character entities. Namespace prefixes are kept verbatim (`w:p`), and
 * lookups match on either the full name or the local name, which is robust
 * for every mainstream producer (Word, LibreOffice, Google Docs export).
 */

export interface XmlElement {
  name: string;
  attrs: Record<string, string>;
  children: XmlElement[];
  /** Concatenated character data directly inside this element. */
  text: string;
  /** Editor placeholder retained in memory but omitted from XML while empty. */
  omitWhenEmpty?: boolean;
}

/** cyrb53 string hash (Bryc, public domain): a fast, well-distributed 53-bit
 * hash returned as a compact decimal string. Used to keep cache keys short. */
export function cyrb53(str: string, seed = 0): string {
  let h1 = 0xdeadbeef ^ seed;
  let h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

export function decodeEntities(s: string): string {
  if (s.indexOf("&") === -1) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|\w+);/g, (m, body: string) => {
    if (body[0] === "#") {
      const code =
        body[1] === "x" || body[1] === "X"
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[body] ?? m;
  });
}

/** Resource caps for parseXml — a denial-of-service guard so a crafted part
 * (millions of empty elements, pathological nesting) cannot exhaust memory on
 * the server or a peer that re-parses a snapshot (plan doc 11, security F4).
 * Generous enough for real documents; override per call if ever needed. */
export interface XmlParseLimits {
  maxLength: number;
  maxNodes: number;
  maxDepth: number;
}
export const DEFAULT_XML_LIMITS: XmlParseLimits = {
  maxLength: 64 * 1024 * 1024, // 64 MB of XML text
  maxNodes: 2_000_000,
  maxDepth: 512,
};

/** The elements an XML part's bytes hold, and how many of them are
 * paragraphs or table cells (local name p or tc), counted without parsing.
 * A start tag inside a comment or CDATA section counts too. */
export function countElements(bytes: Uint8Array): { elements: number; blocks: number } {
  let elements = 0;
  let blocks = 0;
  for (let i = bytes.indexOf(0x3c); i !== -1; i = bytes.indexOf(0x3c, i + 1)) {
    const c = bytes[i + 1];
    if (c === 0x2f /* / */ || c === 0x21 /* ! */ || c === 0x3f /* ? */) continue;
    elements++;
    // The local name: after a colon in the name, if there is one.
    let j = i + 1;
    let local = j;
    while (j < bytes.length && !isNameEnd(bytes[j])) if (bytes[j++] === 0x3a /* : */) local = j;
    const length = j - local;
    if ((length === 1 && bytes[local] === 0x70) /* p */ || (length === 2 && bytes[local] === 0x74 && bytes[local + 1] === 0x63) /* tc */) blocks++;
  }
  return { elements, blocks };
}

/**
 * A copy of `s` that does not point into the string it was cut from. V8
 * makes a substring of 13 or more characters a view into its source, so one
 * long text run kept a document part's whole decoded XML alive after parsing
 * (32 MB for the Yiddish novel's 16 million characters). Concatenating and
 * slicing again copies just this text.
 */
function own(s: string): string {
  return s.length < 13 ? s : (" " + s).slice(1);
}

export function parseXml(input: string, limits: XmlParseLimits = DEFAULT_XML_LIMITS): XmlElement {
  if (input.length > limits.maxLength) {
    throw new Error(`XML parse error: input exceeds ${limits.maxLength} bytes`);
  }
  let i = 0;
  const n = input.length;
  let nodeCount = 0;
  // Strip BOM
  if (input.charCodeAt(0) === 0xfeff) i = 1;

  const root: XmlElement = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XmlElement[] = [root];
  // One string per distinct tag name, attribute name and short attribute
  // value. Each slice of the input is a new string, and a document part
  // repeats the same few hundred of them across half a million elements
  // (w:rPr, w:val, "24"): reusing one copy took a fifth off the parsed tree.
  const strings = new Map<string, string>();
  const intern = (value: string): string => {
    if (value.length > 32) return own(value);
    const known = strings.get(value);
    if (known !== undefined) return known;
    const copy = own(value);
    strings.set(copy, copy);
    return copy;
  };

  const fail = (msg: string): never => {
    throw new Error(`XML parse error at offset ${i}: ${msg}`);
  };

  while (i < n) {
    const lt = input.indexOf("<", i);
    if (lt === -1) break;
    if (lt > i) {
      const text = input.slice(i, lt);
      if (text.trim().length > 0 || stack.length > 1) {
        const top = stack[stack.length - 1];
        if (top !== root) top.text += own(decodeEntities(text));
      }
    }
    i = lt;
    const c1 = input[i + 1];
    if (c1 === "?") {
      const end = input.indexOf("?>", i + 2);
      if (end === -1) fail("unterminated processing instruction");
      i = end + 2;
    } else if (c1 === "!") {
      if (input.startsWith("<!--", i)) {
        const end = input.indexOf("-->", i + 4);
        if (end === -1) fail("unterminated comment");
        i = end + 3;
      } else if (input.startsWith("<![CDATA[", i)) {
        const end = input.indexOf("]]>", i + 9);
        if (end === -1) fail("unterminated CDATA");
        const top = stack[stack.length - 1];
        if (top !== root) top.text += own(input.slice(i + 9, end));
        i = end + 3;
      } else {
        // DOCTYPE or other declaration — skip to next '>'
        const end = input.indexOf(">", i);
        if (end === -1) fail("unterminated declaration");
        i = end + 1;
      }
    } else if (c1 === "/") {
      const end = input.indexOf(">", i + 2);
      if (end === -1) fail("unterminated close tag");
      if (stack.length <= 1) fail("unbalanced close tag");
      // An array grown by push keeps room for more (17 slots for one
      // child); a copy holds exactly its children.
      const closed = stack.pop()!;
      if (closed.children.length > 0) closed.children = closed.children.slice();
      i = end + 1;
    } else {
      // Open tag
      let j = i + 1;
      while (j < n && !isNameEnd(input.charCodeAt(j))) j++;
      const name = intern(input.slice(i + 1, j));
      if (!name) fail("empty tag name");
      const el: XmlElement = { name, attrs: {}, children: [], text: "" };
      // Attributes
      while (j < n) {
        while (j < n && isWhitespace(input.charCodeAt(j))) j++;
        const ch = input[j];
        if (ch === ">" || ch === "/" || ch === undefined) break;
        let k = j;
        while (k < n && input[k] !== "=" && !isWhitespace(input.charCodeAt(k)) && input[k] !== ">" && input[k] !== "/") k++;
        const attrName = intern(input.slice(j, k));
        while (k < n && isWhitespace(input.charCodeAt(k))) k++;
        if (input[k] === "=") {
          k++;
          while (k < n && isWhitespace(input.charCodeAt(k))) k++;
          const quote = input[k];
          if (quote === '"' || quote === "'") {
            const close = input.indexOf(quote, k + 1);
            if (close === -1) fail("unterminated attribute value");
            el.attrs[attrName] = intern(decodeEntities(input.slice(k + 1, close)));
            j = close + 1;
          } else {
            // Unquoted value (not valid XML, but be lenient)
            let m = k;
            while (m < n && !isWhitespace(input.charCodeAt(m)) && input[m] !== ">") m++;
            el.attrs[attrName] = own(decodeEntities(input.slice(k, m)));
            j = m;
          }
        } else {
          if (attrName) el.attrs[attrName] = "";
          j = k;
        }
      }
      const selfClose = input[j] === "/";
      if (selfClose) j++;
      if (input[j] !== ">") fail(`malformed tag <${name}`);
      if (++nodeCount > limits.maxNodes) fail(`too many nodes (> ${limits.maxNodes})`);
      const top = stack[stack.length - 1];
      top.children.push(el);
      if (!selfClose) {
        stack.push(el);
        if (stack.length > limits.maxDepth) fail(`nesting too deep (> ${limits.maxDepth})`);
      }
      i = j + 1;
    }
  }
  if (root.children.length === 0) throw new Error("XML parse error: no root element");
  return root.children[0];
}

function isWhitespace(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}
function isNameEnd(c: number): boolean {
  return isWhitespace(c) || c === 62 /* > */ || c === 47 /* / */;
}

/** Local part of a possibly-prefixed name: `w:p` → `p`. Element names are
 * few and parseXml shares one string for each, so the answer is cached. */
const localNames = new Map<string, string>();
export function localName(name: string): string {
  let local = localNames.get(name);
  if (local === undefined) {
    const idx = name.indexOf(":");
    local = idx === -1 ? name : name.slice(idx + 1);
    if (localNames.size < 10000) localNames.set(name, local);
  }
  return local;
}

export function isEl(el: XmlElement, local: string): boolean {
  return localName(el.name) === local;
}

/** First direct child whose local name matches. */
export function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) if (localName(c.name) === local) return c;
  return undefined;
}

/** All direct children whose local name matches. */
export function children(el: XmlElement | undefined, local: string): XmlElement[] {
  if (!el) return [];
  return el.children.filter((c) => localName(c.name) === local);
}

/** Attribute lookup by local name (`w:val` matches "val"). */
export function attr(el: XmlElement | undefined, local: string): string | undefined {
  if (!el) return undefined;
  if (el.attrs[local] !== undefined) return el.attrs[local];
  for (const key of Object.keys(el.attrs)) {
    if (localName(key) === local) return el.attrs[key];
  }
  return undefined;
}

/** `w:val` of a child element, e.g. childVal(pPr, "jc") → "center". */
export function childVal(el: XmlElement | undefined, local: string): string | undefined {
  return attr(child(el, local), "val");
}

/** Parse an integer attribute; returns undefined when absent or NaN. */
export function intAttr(el: XmlElement | undefined, local: string): number | undefined {
  const v = attr(el, local);
  if (v === undefined) return undefined;
  const num = parseInt(v, 10);
  return Number.isFinite(num) ? num : undefined;
}

/** OOXML on/off value: absent element → undefined; w:val 0/false/off → false; else true. */
export function onOff(el: XmlElement | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, "val");
  if (v === undefined) return true;
  return !(v === "0" || v === "false" || v === "off" || v === "none");
}

// ---------- serialization (for docx write-back) ----------

// One pass, and none at all for the many strings with nothing to escape:
// three chained replaces were a seventh of a large compare's time.
const TEXT_ESCAPES: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;" };
const ATTR_ESCAPES: Record<string, string> = { ...TEXT_ESCAPES, '"': "&quot;", "\n": "&#10;", "\r\n": "&#10;" };
function escapeText(s: string): string {
  return /[&<>]/.test(s) ? s.replace(/[&<>]/g, (c) => TEXT_ESCAPES[c]) : s;
}
function escapeAttr(s: string): string {
  return /[&<>"\n]/.test(s) ? s.replace(/[&<>"]|\r?\n/g, (c) => ATTR_ESCAPES[c]) : s;
}

/**
 * Serialize an element tree back to XML. OOXML elements either carry element
 * children (containers) or character data (w:t, w:instrText) — never meaningful
 * mixed content — so text is emitted only for leaf elements.
 */
export function serializeXml(el: XmlElement, declaration = false): string {
  // Pieces are joined a few thousand at a time: one array of every piece held
  // over five million entries for a 16 MB document part, about 40 MB of
  // pointers alive at once on top of the string they make.
  const chunks: string[] = [];
  const out: string[] = [];
  if (declaration) out.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  const write = (e: XmlElement) => {
    if (out.length > 4096) {
      chunks.push(out.join(""));
      out.length = 0;
    }
    if (
      e.omitWhenEmpty &&
      Object.keys(e.attrs).length === 0 &&
      e.text.length === 0 &&
      e.children.length === 1 &&
      localName(e.children[0].name) === "t" &&
      e.children[0].children.length === 0 &&
      e.children[0].text.length === 0 &&
      Object.keys(e.children[0].attrs).every((key) => key === "xml:space" && e.children[0].attrs[key] === "preserve")
    ) return;
    out.push("<", e.name);
    for (const [k, v] of Object.entries(e.attrs)) {
      out.push(" ", k, '="', escapeAttr(v), '"');
    }
    if (e.children.length === 0 && e.text.length === 0) {
      out.push("/>");
      return;
    }
    out.push(">");
    if (e.children.length > 0) {
      for (const c of e.children) write(c);
    } else {
      out.push(escapeText(e.text));
    }
    out.push("</", e.name, ">");
  };
  write(el);
  if (chunks.length === 0) return out.join("");
  chunks.push(out.join(""));
  return chunks.join("");
}

/** The placeholder serializeXml leaves out: an omitWhenEmpty run holding only an empty w:t. */
export function isEmptyPlaceholder(e: XmlElement): boolean {
  return (
    !!e.omitWhenEmpty &&
    Object.keys(e.attrs).length === 0 &&
    e.text.length === 0 &&
    e.children.length === 1 &&
    localName(e.children[0].name) === "t" &&
    e.children[0].children.length === 0 &&
    e.children[0].text.length === 0 &&
    Object.keys(e.children[0].attrs).every((key) => key === "xml:space" && e.children[0].attrs[key] === "preserve")
  );
}

/**
 * True when two elements write the same XML, attribute order aside, without
 * building either string. compareDocuments asks this of every matched block,
 * and serializing both sides to compare them was a quarter of its time on a
 * document compared with itself.
 */
export function sameXml(a: XmlElement, b: XmlElement): boolean {
  if (a.name !== b.name) return false;
  const ak = Object.keys(a.attrs);
  if (ak.length !== Object.keys(b.attrs).length) return false;
  for (const k of ak) if (a.attrs[k] !== b.attrs[k]) return false;
  // serializeXml writes text only for an element with no children.
  if (a.children.length === 0 && b.children.length === 0) return a.text === b.text;
  const ac = a.children.filter((c) => !isEmptyPlaceholder(c));
  const bc = b.children.filter((c) => !isEmptyPlaceholder(c));
  if (ac.length !== bc.length) return false;
  if (ac.length === 0) return (a.children.length > 0 ? "" : a.text) === (b.children.length > 0 ? "" : b.text);
  for (let i = 0; i < ac.length; i++) if (!sameXml(ac[i], bc[i])) return false;
  return true;
}

/** Deep-clone an element (attrs and children copied, text preserved). */
export function cloneXml(el: XmlElement): XmlElement {
  return {
    name: el.name,
    attrs: { ...el.attrs },
    text: el.text,
    children: el.children.map(cloneXml),
    ...(el.omitWhenEmpty ? { omitWhenEmpty: true } : {}),
  };
}

/** Descend through a path of local names. */
export function path(el: XmlElement | undefined, ...locals: string[]): XmlElement | undefined {
  let cur = el;
  for (const l of locals) {
    cur = child(cur, l);
    if (!cur) return undefined;
  }
  return cur;
}
