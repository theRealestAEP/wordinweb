import { DocxDocument } from "../docx.js";
import { MathNode } from "../model.js";
import { XmlElement, cloneXml, localName } from "../xml.js";

/**
 * Math editing: equations round-trip through a linear form ("e^x = 1+x+x/2",
 * groups in braces: {x+1}/{2y}, roots as √{…}) that users can edit in place;
 * the result is written back as OMML.
 */

/** MathNode AST -> editable linear text. */
export function linearizeMath(nodes: MathNode[]): string {
  const group = (inner: MathNode[]): string => {
    const s = linearizeMath(inner);
    return s.length === 1 ? s : `{${s}}`;
  };
  let out = "";
  for (const n of nodes) {
    switch (n.t) {
      case "run":
        out += n.text;
        break;
      case "sup":
      case "sub":
        out += linearizeMath(n.base) + (n.t === "sup" ? "^" : "_") + group(n.script);
        break;
      case "frac":
        out += group(n.num) + "/" + group(n.den);
        break;
      case "rad": {
        // Degree (nth root) linearizes as √[deg]{e} — the index stays plain
        // editable text (∛ would bake the 3 into one atomic character). The
        // parser still ACCEPTS ∛/∜ as input shorthands.
        const deg = n.deg && n.deg.length ? linearizeMath(n.deg) : "";
        out += deg ? "√[" + deg + "]" + group(n.e) : "√" + group(n.e);
        break;
      }
      case "nary":
        out += n.chr + (n.sub.length ? "_" + group(n.sub) : "") + (n.sup.length ? "^" + group(n.sup) : "") + group(n.e);
        break;
      case "dlm":
        out += n.beg + n.e.map((part) => linearizeMath(part)).join("|") + n.end;
        break;
      case "mat":
        out += "[" + n.rows.map((row) => row.map((cell) => linearizeMath(cell)).join("&")).join(";") + "]";
        break;
    }
  }
  return out;
}

// N-ary operators (∑ ∏ ∫ …) linearize as "chr(_sub)(^sup)integrand"; the
// integrand is exactly one unit so the parser can hand it back structurally.
const NARY_CHRS = new Set([
  "∑", "∏", "∐", "∫", "∬", "∭", "∮", "∯",
  "∰", "⋀", "⋁", "⋂", "⋃", "⨀", "⨁", "⨂",
  "⨄", "⨆",
]);
// Delimiter open -> close. "{" stays a grouping brace, never a delimiter.
const DELIM_CLOSE: Record<string, string> = {
  "(": ")", "[": "]", "⟨": "⟩", "⌊": "⌋", "⌈": "⌉",
  "|": "|", "‖": "‖",
};
const DELIM_OPENS = "([⟨⌊⌈|‖";
const DELIM_CLOSES = ")]⟩⌋⌉|‖";

/** Split at top-level occurrences of any separator char (ignores {} and
 * bracketed depth), so "a|b" splits but "(a|b)" and "{a|b}" do not. */
function splitTop(text: string, seps: string): string[] {
  const parts: string[] = [];
  let cur = "";
  let depth = 0;
  for (let k = 0; k < text.length; k++) {
    const c = text[k];
    if (c === "{" || (DELIM_OPENS.includes(c) && c !== "|" && c !== "‖")) depth++;
    else if (c === "}" || (DELIM_CLOSES.includes(c) && c !== "|" && c !== "‖")) depth = Math.max(0, depth - 1);
    else if (depth === 0 && seps.includes(c)) {
      parts.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  parts.push(cur);
  return parts;
}

/** Editable linear text -> MathNode AST. Scripts bind tighter than "/". */
export function parseMathLinear(input: string): MathNode[] {
  let i = 0;

  const parseGroup = (): MathNode[] => {
    // at "{": consume the balanced group
    i++; // {
    const start = i;
    let depth = 1;
    while (i < input.length && depth > 0) {
      if (input[i] === "{") depth++;
      else if (input[i] === "}") depth--;
      if (depth > 0) i++;
    }
    const inner = input.slice(start, i);
    i++; // }
    return parseMathLinear(inner);
  };

  // A balanced delimited region: strip the outer pair, tracking {} and nested
  // like-delimiters so "((a))" and "(a{b})" close at the right bracket.
  const scanDelim = (open: string, close: string): string => {
    i++; // opener
    const start = i;
    let brace = 0;
    let nest = 0;
    while (i < input.length) {
      const c = input[i];
      if (c === "{") brace++;
      else if (c === "}") brace = Math.max(0, brace - 1);
      else if (brace === 0 && open !== close && c === open) nest++;
      else if (brace === 0 && c === close) {
        if (nest === 0) break;
        nest--;
      }
      i++;
    }
    const inner = input.slice(start, i);
    if (input[i] === close) i++; // closer
    return inner;
  };

  const parseDelim = (open: string): MathNode[] => {
    const close = DELIM_CLOSE[open];
    const inner = scanDelim(open, close);
    // "[…]" carrying top-level & or ; is a matrix; otherwise a bracket group.
    if (open === "[" && (splitTop(inner, "&").length > 1 || splitTop(inner, ";").length > 1)) {
      const rows = splitTop(inner, ";").map((row) => splitTop(row, "&").map((cell) => parseMathLinear(cell)));
      return [{ t: "mat", rows }];
    }
    const parts = splitTop(inner, "|").map((part) => parseMathLinear(part));
    return [{ t: "dlm", beg: open, end: close, e: parts }];
  };

  const parseNary = (): MathNode[] => {
    const chr = input[i];
    i++;
    let sub: MathNode[] = [];
    let sup: MathNode[] = [];
    const readScript = (): MathNode[] => (input[i] === "{" ? parseGroup() : i < input.length ? parseUnit() : []);
    for (let g = 0; g < 2; g++) {
      if (input[i] === "_" && sub.length === 0) {
        i++;
        sub = readScript();
      } else if (input[i] === "^" && sup.length === 0) {
        i++;
        sup = readScript();
      }
    }
    const startsIntegrand = i < input.length && input[i] !== "^" && input[i] !== "_" && input[i] !== "/";
    const e = startsIntegrand ? (input[i] === "{" ? parseGroup() : parseUnit()) : [];
    return [{ t: "nary", chr, sub, sup, e }];
  };

  /** One unit: a group, root, n-ary, delimiter/matrix, or a single run char. */
  const parseUnit = (): MathNode[] => {
    const ch = input[i];
    if (ch === "{") return parseGroup();
    if (ch === "√" || ch === "∛" || ch === "∜") {
      i++;
      let deg: MathNode[] =
        ch === "∛" ? [{ t: "run", text: "3" }] : ch === "∜" ? [{ t: "run", text: "4" }] : [];
      // "√[deg]{e}": a bracket group right after √ is the root's index.
      if (ch === "√" && input[i] === "[") deg = parseMathLinear(scanDelim("[", "]"));
      const e = input[i] === "{" ? parseGroup() : parseUnit();
      return deg.length ? [{ t: "rad", e, deg }] : [{ t: "rad", e }];
    }
    if (NARY_CHRS.has(ch)) return parseNary();
    if (ch in DELIM_CLOSE) return parseDelim(ch);
    i++;
    return [{ t: "run", text: ch }];
  };

  const out: MathNode[] = [];
  const push = (nodes: MathNode[]) => {
    for (const n of nodes) {
      const last = out[out.length - 1];
      if (n.t === "run" && last && last.t === "run") last.text += n.text;
      else out.push(n);
    }
  };

  // Each unit stays atomic while trailing operators (^ _ /) consume it, so
  // "{a+b}/{2c}" fractions the whole group, not its last character.
  while (i < input.length) {
    let unit = parseUnit();
    while (i < input.length && (input[i] === "^" || input[i] === "_" || input[i] === "/")) {
      const op = input[i];
      i++;
      const arg = input[i] === "{" ? parseGroup() : parseUnit();
      if (op === "/") unit = [{ t: "frac", num: unit, den: arg }];
      else unit = [{ t: op === "^" ? "sup" : "sub", base: unit, script: arg }];
    }
    push(unit);
  }
  return out;
}

/** Rewrite an m:oMath element's content from linear text. */
export function setMathLinear(doc: DocxDocument, oMathEl: XmlElement, text: string): boolean {
  const nodes = parseMathLinear(text.trim());
  if (nodes.length === 0) return false;
  const m = oMathEl.name.includes(":") ? oMathEl.name.slice(0, oMathEl.name.indexOf(":") + 1) : "m:";
  oMathEl.children = nodes.map((n) => buildOmml(n, m));
  doc.refresh();
  return true;
}

/** Remove one complete OMML equation from the document. */
export function deleteMath(doc: DocxDocument, oMathEl: XmlElement): boolean {
  const parent = doc.findParentOf(oMathEl);
  if (!parent) return false;

  let container = parent;
  let target = oMathEl;
  if (localName(parent.name) === "oMathPara") {
    const equations = parent.children.filter((child) => localName(child.name) === "oMath");
    if (equations.length === 1) {
      const grandparent = doc.findParentOf(parent);
      if (!grandparent) return false;
      container = grandparent;
      target = parent;
    }
  }

  const index = container.children.indexOf(target);
  if (index < 0) return false;
  container.children.splice(index, 1);
  doc.refresh();
  return true;
}

/** Insert a new inline OMML equation at a text position. */
export function insertMathAt(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  text: string,
): XmlElement | null {
  const nodes = parseMathLinear(text.trim());
  if (nodes.length === 0) return null;
  const run = doc.findParentOf(t);
  const parent = run && doc.findParentOf(run);
  if (!run || !parent || localName(run.name) !== "r") return null;

  const contains = (root: XmlElement): boolean =>
    root === t || root.children.some(contains);
  const root = doc.editableRoots().find(contains);
  if (root && !Object.prototype.hasOwnProperty.call(root.attrs, "xmlns:m")) {
    root.attrs["xmlns:m"] = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  }
  const equation = el("m:oMath", nodes.map((node) => buildOmml(node, "m:")));
  const runIndex = parent.children.indexOf(run);
  const textIndex = run.children.indexOf(t);
  if (runIndex < 0 || textIndex < 0 || !root) return null;
  const at = Math.max(0, Math.min(offset, t.text.length));
  const rPr = run.children.find((child) => localName(child.name) === "rPr");
  const makeText = (text: string): XmlElement => ({
    name: t.name,
    attrs: { ...t.attrs, "xml:space": "preserve" },
    children: [],
    text,
  });
  const makeRun = (content: XmlElement[]): XmlElement => ({
    name: run.name,
    attrs: { ...run.attrs },
    children: [...(rPr ? [cloneXml(rPr)] : []), ...content],
    text: "",
  });
  const before = run.children.slice(0, textIndex).filter((child) => localName(child.name) !== "rPr");
  const after = run.children.slice(textIndex + 1).filter((child) => localName(child.name) !== "rPr");
  let beforeRun: XmlElement | null;
  let afterRun: XmlElement | null;
  if (at === 0) {
    beforeRun = before.length > 0 ? makeRun(before) : null;
    run.children = [...(rPr ? [rPr] : []), t, ...after];
    afterRun = run;
  } else {
    const tail = at < t.text.length ? makeText(t.text.slice(at)) : null;
    t.text = t.text.slice(0, at);
    run.children = [...(rPr ? [rPr] : []), ...before, t];
    beforeRun = run;
    afterRun = tail || after.length > 0 ? makeRun([...(tail ? [tail] : []), ...after]) : null;
  }
  parent.children.splice(
    runIndex,
    1,
    ...(beforeRun ? [beforeRun] : []),
    equation,
    ...(afterRun ? [afterRun] : []),
  );
  doc.refresh();
  return equation;
}

function el(name: string, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs: {}, children, text };
}

function buildOmml(node: MathNode, m: string): XmlElement {
  switch (node.t) {
    case "run":
      return el(`${m}r`, [el(`${m}t`, [], node.text)]);
    case "sup":
    case "sub": {
      const tag = node.t === "sup" ? "sSup" : "sSub";
      const scriptTag = node.t === "sup" ? "sup" : "sub";
      return el(`${m}${tag}`, [
        el(`${m}e`, node.base.map((n) => buildOmml(n, m))),
        el(`${m}${scriptTag}`, node.script.map((n) => buildOmml(n, m))),
      ]);
    }
    case "frac":
      return el(`${m}f`, [
        el(`${m}num`, node.num.map((n) => buildOmml(n, m))),
        el(`${m}den`, node.den.map((n) => buildOmml(n, m))),
      ]);
    case "rad": {
      const kids: XmlElement[] = [];
      if (node.deg && node.deg.length) {
        kids.push(el(`${m}deg`, node.deg.map((n) => buildOmml(n, m))));
      } else {
        const hide = el(`${m}degHide`);
        hide.attrs[`${m}val`] = "1";
        kids.push(el(`${m}radPr`, [hide]), el(`${m}deg`));
      }
      kids.push(el(`${m}e`, node.e.map((n) => buildOmml(n, m))));
      return el(`${m}rad`, kids);
    }
    case "nary": {
      const pr = el(`${m}naryPr`, [el(`${m}chr`)]);
      pr.children[0].attrs[`${m}val`] = node.chr;
      return el(`${m}nary`, [
        pr,
        el(`${m}sub`, node.sub.map((n) => buildOmml(n, m))),
        el(`${m}sup`, node.sup.map((n) => buildOmml(n, m))),
        el(`${m}e`, node.e.map((n) => buildOmml(n, m))),
      ]);
    }
    case "dlm": {
      const beg = el(`${m}begChr`);
      beg.attrs[`${m}val`] = node.beg;
      const end = el(`${m}endChr`);
      end.attrs[`${m}val`] = node.end;
      return el(`${m}d`, [el(`${m}dPr`, [beg, end]), ...node.e.map((part) => el(`${m}e`, part.map((n) => buildOmml(n, m))))]);
    }
    case "mat":
      return el(`${m}m`, node.rows.map((row) => el(`${m}mr`, row.map((cell) => el(`${m}e`, cell.map((n) => buildOmml(n, m)))))));
    case "eqarr":
      return el(`${m}eqArr`, node.rows.map((row) => el(`${m}e`, row.map((n) => buildOmml(n, m)))));
    case "acc": {
      const chr = el(`${m}chr`);
      chr.attrs[`${m}val`] = node.chr;
      return el(`${m}acc`, [el(`${m}accPr`, [chr]), el(`${m}e`, node.e.map((n) => buildOmml(n, m)))]);
    }
    case "grp": {
      const chr = el(`${m}chr`);
      chr.attrs[`${m}val`] = node.chr;
      const pos = el(`${m}pos`);
      pos.attrs[`${m}val`] = node.pos;
      const vjc = el(`${m}vertJc`);
      vjc.attrs[`${m}val`] = node.vertJc;
      return el(`${m}groupChr`, [el(`${m}groupChrPr`, [chr, pos, vjc]), el(`${m}e`, node.e.map((n) => buildOmml(n, m)))]);
    }
    case "lim":
      return el(`${m}${node.pos === "low" ? "limLow" : "limUpp"}`, [
        el(`${m}e`, node.e.map((n) => buildOmml(n, m))),
        el(`${m}lim`, node.lim.map((n) => buildOmml(n, m))),
      ]);
  }
}

/**
 * Move an equation to a text position: detach the m:oMath from its
 * paragraph and re-insert it at `offset` inside the w:t `t` (splitting the
 * destination run when the drop lands mid-text).
 */
export function moveMath(doc: DocxDocument, oMathEl: XmlElement, t: XmlElement, offset: number): boolean {
  const curParent = doc.findParentOf(oMathEl);
  const rEl = doc.findParentOf(t);
  const pEl = rEl && doc.findParentOf(rEl);
  if (!curParent || !rEl || !pEl || localName(rEl.name) !== "r") return false;
  if (rEl === oMathEl || curParent === oMathEl) return false;
  curParent.children.splice(curParent.children.indexOf(oMathEl), 1);
  const rw = rEl.name.includes(":") ? rEl.name.slice(0, rEl.name.indexOf(":") + 1) : "";
  const rIdx = pEl.children.indexOf(rEl);
  if (offset >= t.text.length) {
    pEl.children.splice(rIdx + 1, 0, oMathEl);
  } else if (offset <= 0) {
    pEl.children.splice(rIdx, 0, oMathEl);
  } else {
    const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
    const clone = (e: XmlElement): XmlElement => ({ name: e.name, attrs: { ...e.attrs }, children: e.children.map(clone), text: e.text });
    const tailT: XmlElement = { name: `${rw}t`, attrs: { "xml:space": "preserve" }, children: [], text: t.text.slice(offset) };
    t.text = t.text.slice(0, offset);
    const tail: XmlElement = { name: `${rw}r`, attrs: {}, children: [...(rPr ? [clone(rPr)] : []), tailT], text: "" };
    pEl.children.splice(rIdx + 1, 0, oMathEl, tail);
  }
  doc.refresh();
  return true;
}

/** Derive the math AST from a live oMath element (reflects current XML). */
function ommlToNodes(e: XmlElement): MathNode[] {
  const ln = localName(e.name);
  const kids = (name: string): MathNode[] => {
    const c = e.children.find((ch) => localName(ch.name) === name);
    return c ? ommlToNodes(c) : [];
  };
  const chrAttr = (prName: string, chrName: string, dflt: string): string => {
    const pr = e.children.find((c) => localName(c.name) === prName);
    const c = pr?.children.find((ch) => localName(ch.name) === chrName);
    const k = c && Object.keys(c.attrs).find((key) => localName(key) === "val");
    return c && k ? c.attrs[k] : dflt;
  };
  if (ln === "f") return [{ t: "frac", num: kids("num"), den: kids("den") }];
  if (ln === "nary") {
    return [{ t: "nary", chr: chrAttr("naryPr", "chr", "\u222b"), sub: kids("sub"), sup: kids("sup"), e: kids("e") }];
  }
  if (ln === "d") {
    const pr = e.children.find((c) => localName(c.name) === "dPr");
    const chr = (name: string, dflt: string) => {
      const c = pr?.children.find((ch) => localName(ch.name) === name);
      const k = c && Object.keys(c.attrs).find((key) => localName(key) === "val");
      return c && k ? c.attrs[k] : dflt;
    };
    const parts = e.children.filter((c) => localName(c.name) === "e").map(ommlToNodes);
    return [{ t: "dlm", beg: chr("begChr", "("), end: chr("endChr", ")"), e: parts }];
  }
  if (ln === "m" && e.children.some((c) => localName(c.name) === "mr")) {
    const rows = e.children
      .filter((c) => localName(c.name) === "mr")
      .map((mr) => mr.children.filter((c) => localName(c.name) === "e").map(ommlToNodes));
    return [{ t: "mat", rows }];
  }
  // Constructs with no faithful linear syntax: keep them as their own node so
  // the round-trip check (isLinearSafe) sees the mismatch and refuses to let
  // text editing silently drop the accent/limit/over-under/array structure.
  if (ln === "acc") return [{ t: "acc", chr: chrAttr("accPr", "chr", "\u0302"), e: kids("e") }];
  if (ln === "groupChr") {
    return [
      {
        t: "grp",
        chr: chrAttr("groupChrPr", "chr", "\u23df"),
        pos: chrAttr("groupChrPr", "pos", "bot") === "top" ? "top" : "bot",
        vertJc: chrAttr("groupChrPr", "vertJc", "bot") === "top" ? "top" : "bot",
        e: kids("e"),
      },
    ];
  }
  if (ln === "limLow" || ln === "limUpp") {
    return [{ t: "lim", pos: ln === "limLow" ? "low" : "upp", e: kids("e"), lim: kids("lim") }];
  }
  if (ln === "eqArr") {
    const rows = e.children.filter((c) => localName(c.name) === "e").map(ommlToNodes);
    return [{ t: "eqarr", rows }];
  }
  if (ln === "sSup") return [{ t: "sup", base: kids("e"), script: kids("sup") }];
  if (ln === "sSub") return [{ t: "sub", base: kids("e"), script: kids("sub") }];
  if (ln === "rad") {
    const hide = chrAttr("radPr", "degHide", "0");
    const deg = hide === "1" || hide === "true" || hide === "on" ? [] : kids("deg");
    return deg.length ? [{ t: "rad", e: kids("e"), deg }] : [{ t: "rad", e: kids("e") }];
  }
  if (ln === "t") return e.text ? [{ t: "run", text: e.text }] : [];
  const out: MathNode[] = [];
  for (const c of e.children) {
    for (const n of ommlToNodes(c)) {
      const last = out[out.length - 1];
      if (n.t === "run" && last && last.t === "run") last.text += n.text;
      else out.push(n);
    }
  }
  return out;
}

/** The math AST currently in an oMath element (for prefilling the editor). */
export function mathLinearOf(doc: DocxDocument, oMathEl: XmlElement): string {
  void doc;
  return linearizeMath(ommlToNodes(oMathEl));
}

/** Two ASTs are structurally identical (same node kinds, nesting, and the
 * chr/beg/end/pos discriminants) \u2014 runs compare by presence, not text. */
function sameStructure(a: MathNode[], b: MathNode[]): boolean {
  if (a.length !== b.length) return false;
  for (let k = 0; k < a.length; k++) {
    const x = a[k];
    const y = b[k];
    if (x.t !== y.t) return false;
    switch (x.t) {
      case "run":
        break;
      case "sup":
      case "sub":
        if (!sameStructure(x.base, (y as typeof x).base) || !sameStructure(x.script, (y as typeof x).script)) return false;
        break;
      case "frac":
        if (!sameStructure(x.num, (y as typeof x).num) || !sameStructure(x.den, (y as typeof x).den)) return false;
        break;
      case "rad": {
        const yy = y as typeof x;
        if (!sameStructure(x.e, yy.e) || !sameStructure(x.deg ?? [], yy.deg ?? [])) return false;
        break;
      }
      case "nary": {
        const yy = y as typeof x;
        if (x.chr !== yy.chr || !sameStructure(x.sub, yy.sub) || !sameStructure(x.sup, yy.sup) || !sameStructure(x.e, yy.e))
          return false;
        break;
      }
      case "dlm": {
        const yy = y as typeof x;
        if (x.beg !== yy.beg || x.end !== yy.end || x.e.length !== yy.e.length) return false;
        for (let p = 0; p < x.e.length; p++) if (!sameStructure(x.e[p], yy.e[p])) return false;
        break;
      }
      case "mat": {
        const yy = y as typeof x;
        if (x.rows.length !== yy.rows.length) return false;
        for (let r = 0; r < x.rows.length; r++) {
          if (x.rows[r].length !== yy.rows[r].length) return false;
          for (let c = 0; c < x.rows[r].length; c++) if (!sameStructure(x.rows[r][c], yy.rows[r][c])) return false;
        }
        break;
      }
      default:
        // acc / grp / lim / eqarr have no linear syntax: never round-trip-safe.
        return false;
    }
  }
  return true;
}

/**
 * True when the equation round-trips losslessly through the linear text form,
 * so free-text editing can never silently change its structure. Equations that
 * fail this (n-ary/delimiter/matrix the parser can't rebuild, accents, limits,
 * over/under groups, equation arrays) are shown read-only instead.
 */
export function isLinearSafe(oMathEl: XmlElement): boolean {
  const nodes = ommlToNodes(oMathEl);
  const text = linearizeMath(nodes);
  if (!text.trim()) return false;
  return sameStructure(nodes, parseMathLinear(text));
}
