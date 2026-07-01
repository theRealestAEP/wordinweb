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

export function parseXml(input: string): XmlElement {
  let i = 0;
  const n = input.length;
  // Strip BOM
  if (input.charCodeAt(0) === 0xfeff) i = 1;

  const root: XmlElement = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XmlElement[] = [root];

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
        if (top !== root) top.text += decodeEntities(text);
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
        if (top !== root) top.text += input.slice(i + 9, end);
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
      stack.pop();
      i = end + 1;
    } else {
      // Open tag
      let j = i + 1;
      while (j < n && !isNameEnd(input.charCodeAt(j))) j++;
      const name = input.slice(i + 1, j);
      if (!name) fail("empty tag name");
      const el: XmlElement = { name, attrs: {}, children: [], text: "" };
      // Attributes
      while (j < n) {
        while (j < n && isWhitespace(input.charCodeAt(j))) j++;
        const ch = input[j];
        if (ch === ">" || ch === "/" || ch === undefined) break;
        let k = j;
        while (k < n && input[k] !== "=" && !isWhitespace(input.charCodeAt(k)) && input[k] !== ">" && input[k] !== "/") k++;
        const attrName = input.slice(j, k);
        while (k < n && isWhitespace(input.charCodeAt(k))) k++;
        if (input[k] === "=") {
          k++;
          while (k < n && isWhitespace(input.charCodeAt(k))) k++;
          const quote = input[k];
          if (quote === '"' || quote === "'") {
            const close = input.indexOf(quote, k + 1);
            if (close === -1) fail("unterminated attribute value");
            el.attrs[attrName] = decodeEntities(input.slice(k + 1, close));
            j = close + 1;
          } else {
            // Unquoted value (not valid XML, but be lenient)
            let m = k;
            while (m < n && !isWhitespace(input.charCodeAt(m)) && input[m] !== ">") m++;
            el.attrs[attrName] = decodeEntities(input.slice(k, m));
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
      const top = stack[stack.length - 1];
      top.children.push(el);
      if (!selfClose) stack.push(el);
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

/** Local part of a possibly-prefixed name: `w:p` → `p`. */
export function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx === -1 ? name : name.slice(idx + 1);
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

/** Descend through a path of local names. */
export function path(el: XmlElement | undefined, ...locals: string[]): XmlElement | undefined {
  let cur = el;
  for (const l of locals) {
    cur = child(cur, l);
    if (!cur) return undefined;
  }
  return cur;
}
