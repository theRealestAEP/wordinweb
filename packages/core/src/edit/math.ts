import { DocxDocument } from "../docx.js";
import { MathNode } from "../model.js";
import { XmlElement, cloneXml, localName, serializeXml } from "../xml.js";

/**
 * Math editing: equations round-trip through a linear form ("e^x = 1+x+x/2",
 * groups in braces: {x+1}/{2y}, roots as √{…}) that users can edit in place;
 * the result is written back as OMML.
 *
 * The linear form spells out an equation's STRUCTURE and its text, never its
 * formatting. Everything else — run fonts, m:ctrlPr, an n-ary's m:limLoc, a
 * matrix's column specs — rides along on each node's source element and is
 * re-emitted verbatim (see buildOmml), so editing the text of an equation
 * cannot strip the properties the text does not mention.
 */

/**
 * Structure glyphs. Word's own linear format names a matrix ■(a&b@c&d) and an
 * equation array █(a@b); we keep our bracket-and-semicolon rows and borrow
 * only the two leading glyphs, so strings written against the older grammar
 * still mean what they meant.
 */
const MATRIX_MARK = "■";
const EQARR_MARK = "█";
/** m:acc combining marks trail their base, exactly as they do in plain text. */
const ACCENT_RE = /[̀-ͯ⃐-⃰]/;
/** m:groupChr over/under characters lead their argument. */
const GROUP_CHRS = new Set(["⏜", "⏝", "⏞", "⏟", "⏠", "⏡", "⎴", "⎵"]);
/** The three that sit ABOVE the group; the rest sit below. */
const GROUP_TOP_CHRS = new Set(["⏜", "⏞", "⏠", "⎴"]);
/** m:limUpp / m:limLow stack a limit over/under an operator, as Word does. */
const LIM_UPP = "┬";
const LIM_LOW = "┴";
/** Word's glyph for a delimiter with nothing on its closing side ("cases"). */
const CASES_END = "┤";

/** MathNode AST -> editable linear text. */
export function linearizeMath(nodes: MathNode[]): string {
  return lin(nodes, false);
}

/**
 * `inCell` says the text sits inside a matrix or equation-array cell, where &
 * and ; end the cell and the row. Outside one they are ordinary characters and
 * stay unescaped, which keeps "F(α;β)" readable.
 */
function lin(nodes: MathNode[], inCell: boolean): string {
  const group = (inner: MathNode[]): string => {
    const s = lin(inner, inCell);
    return s.length === 1 ? s : `{${s}}`;
  };
  /**
   * A script binds to ONE unit. One character, or one node that is not a run,
   * is already that unit — the parser folds ^ and _ onto whatever it last
   * built. Anything longer has to be braced, or "xy^2" would come back with
   * the 2 riding on the y alone.
   */
  const scriptBase = (base: MathNode[]): string => {
    const s = lin(base, inCell);
    return s.length === 1 || (base.length === 1 && base[0].t !== "run") ? s : `{${s}}`;
  };
  let out = "";
  for (const n of nodes) {
    switch (n.t) {
      case "run":
        out += escapeRun(n.text, inCell);
        break;
      case "sup":
      case "sub":
        out += scriptBase(n.base) + (n.t === "sup" ? "^" : "_") + group(n.script);
        break;
      case "frac":
        out += group(n.num) + "/" + group(n.den);
        break;
      case "rad": {
        // Degree (nth root) linearizes as √[deg]{e} — the index stays plain
        // editable text (∛ would bake the 3 into one atomic character). The
        // parser still ACCEPTS ∛/∜ as input shorthands.
        const deg = n.deg && n.deg.length ? lin(n.deg, inCell) : "";
        out += deg ? "√[" + deg + "]" + group(n.e) : "√" + group(n.e);
        break;
      }
      case "nary":
        out += n.chr + (n.sub.length ? "_" + group(n.sub) : "") + (n.sup.length ? "^" + group(n.sup) : "") + group(n.e);
        break;
      case "dlm": {
        const parts = n.e.map((part) => lin(part, inCell)).join("|");
        // Word's "cases" bracket — a brace on the left and nothing on the
        // right — takes ┤, its own glyph for a missing closing delimiter,
        // because a bare { already opens a group here.
        out += n.beg === "{" && n.end === "" ? `{${parts}}${CASES_END}` : n.beg + parts + n.end;
        break;
      }
      case "mat": {
        const body = n.rows.map((row) => row.map((cell) => lin(cell, true)).join("&")).join(";");
        // A one-cell matrix carries no & or ; to tell it apart from a bracket
        // group, so it takes the explicit marker. Bigger ones stay bare.
        const bare = n.rows.length > 1 || (n.rows[0]?.length ?? 0) > 1;
        out += (bare ? "" : MATRIX_MARK) + "[" + body + "]";
        break;
      }
      case "eqarr":
        out += EQARR_MARK + "[" + n.rows.map((row) => lin(row, true)).join(";") + "]";
        break;
      case "acc":
        out += group(n.e) + n.chr;
        break;
      case "grp":
        out += n.chr + group(n.e);
        break;
      case "lim":
        out += group(n.e) + (n.pos === "low" ? LIM_LOW : LIM_UPP) + group(n.lim);
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

/**
 * Every character the parser reads as structure. Inside run text each one
 * takes a leading backslash, so an equation whose author simply TYPED "(x)" or
 * "-1/2" comes back as those characters and not as a delimiter or a fraction.
 * The backslash escapes itself.
 */
const SPECIAL = new Set<string>([
  "\\", "{", "}", "^", "_", "/", LIM_UPP, LIM_LOW, CASES_END, "√", "∛", "∜", MATRIX_MARK, EQARR_MARK,
  ...DELIM_OPENS, ...DELIM_CLOSES, ...NARY_CHRS, ...GROUP_CHRS,
]);

function escapeRun(text: string, inCell: boolean): string {
  let out = "";
  for (const ch of text) {
    if (SPECIAL.has(ch) || ACCENT_RE.test(ch) || (inCell && (ch === "&" || ch === ";"))) out += "\\";
    out += ch;
  }
  return out;
}

/** Split at top-level occurrences of any separator char (ignores {} and
 * bracketed depth), so "a|b" splits but "(a|b)" and "{a|b}" do not. */
function splitTop(text: string, seps: string): string[] {
  const parts: string[] = [];
  let cur = "";
  let depth = 0;
  for (let k = 0; k < text.length; k++) {
    const c = text[k];
    if (c === "\\" && k + 1 < text.length) {
      cur += c + text[k + 1];
      k++;
      continue;
    }
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

  /** At "{": consume the balanced group and hand back its inner text. */
  const scanGroup = (): string => {
    i++; // {
    const start = i;
    let depth = 1;
    while (i < input.length && depth > 0) {
      if (input[i] === "\\" && i + 1 < input.length) i++;
      else if (input[i] === "{") depth++;
      else if (input[i] === "}") depth--;
      if (depth > 0) i++;
    }
    const inner = input.slice(start, i);
    i++; // }
    return inner;
  };

  const parseGroup = (): MathNode[] => parseMathLinear(scanGroup());

  // A balanced delimited region: strip the outer pair, tracking {} and nested
  // like-delimiters so "((a))" and "(a{b})" close at the right bracket.
  const scanDelim = (open: string, close: string): string => {
    i++; // opener
    const start = i;
    let brace = 0;
    let nest = 0;
    while (i < input.length) {
      const c = input[i];
      if (c === "\\" && i + 1 < input.length) i++;
      else if (c === "{") brace++;
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
    const readScript = (): MathNode[] => (input[i] === "{" ? parseGroup() : i < input.length ? parseAtom() : []);
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
    const e = startsIntegrand ? (input[i] === "{" ? parseGroup() : parseAtom()) : [];
    return [{ t: "nary", chr, sub, sup, e }];
  };

  /** One unit: a group, root, n-ary, delimiter/matrix, or a single run char. */
  const parseUnit = (): MathNode[] => {
    const ch = input[i];
    // "\(" and friends: the next character is literal text, whatever it means
    // to the grammar.
    if (ch === "\\" && i + 1 < input.length) {
      i += 2;
      return [{ t: "run", text: input[i - 1] }];
    }
    if (ch === "{") {
      const inner = scanGroup();
      // "{…}┤" is the cases bracket, not a plain group.
      if (input[i] === CASES_END) {
        i++;
        return [{ t: "dlm", beg: "{", end: "", e: splitTop(inner, "|").map((part) => parseMathLinear(part)) }];
      }
      return parseMathLinear(inner);
    }
    if (ch === "√" || ch === "∛" || ch === "∜") {
      i++;
      let deg: MathNode[] =
        ch === "∛" ? [{ t: "run", text: "3" }] : ch === "∜" ? [{ t: "run", text: "4" }] : [];
      // "√[deg]{e}": a bracket group right after √ is the root's index.
      if (ch === "√" && input[i] === "[") deg = parseMathLinear(scanDelim("[", "]"));
      const e = input[i] === "{" ? parseGroup() : parseAtom();
      return deg.length ? [{ t: "rad", e, deg }] : [{ t: "rad", e }];
    }
    // "■[…]" forces a matrix even with one cell; "█[a;b]" is an equation array.
    if ((ch === MATRIX_MARK || ch === EQARR_MARK) && input[i + 1] === "[") {
      i++;
      const inner = scanDelim("[", "]");
      if (ch === EQARR_MARK) return [{ t: "eqarr", rows: splitTop(inner, ";").map((row) => parseMathLinear(row)) }];
      return [{ t: "mat", rows: splitTop(inner, ";").map((row) => splitTop(row, "&").map((cell) => parseMathLinear(cell))) }];
    }
    if (GROUP_CHRS.has(ch)) {
      i++;
      const e = input[i] === "{" ? parseGroup() : parseAtom();
      // The character says which side it sits on; the content takes the other.
      const pos = GROUP_TOP_CHRS.has(ch) ? "top" : "bot";
      return [{ t: "grp", chr: ch, pos, vertJc: pos === "top" ? "bot" : "top", e }];
    }
    if (NARY_CHRS.has(ch)) return parseNary();
    if (ch in DELIM_CLOSE) return parseDelim(ch);
    i++;
    return [{ t: "run", text: ch }];
  };

  /** A unit plus the combining accents that trail it: "x̂", "{a+b}⃗". */
  const parseAtom = (): MathNode[] => {
    let e = parseUnit();
    while (i < input.length && ACCENT_RE.test(input[i])) {
      e = [{ t: "acc", chr: input[i], e }];
      i++;
    }
    return e;
  };

  const out: MathNode[] = [];
  const push = (nodes: MathNode[]) => {
    for (const n of nodes) {
      const last = out[out.length - 1];
      if (n.t === "run" && last && last.t === "run") last.text += n.text;
      else out.push(n);
    }
  };

  // Each unit stays atomic while trailing operators (^ _ / ┬ ┴) consume it, so
  // "{a+b}/{2c}" fractions the whole group, not its last character.
  const OPS = "^_/" + LIM_UPP + LIM_LOW;
  while (i < input.length) {
    let unit = parseAtom();
    while (i < input.length && OPS.includes(input[i])) {
      const op = input[i];
      i++;
      const arg = input[i] === "{" ? parseGroup() : parseAtom();
      if (op === "/") unit = [{ t: "frac", num: unit, den: arg }];
      else if (op === LIM_UPP || op === LIM_LOW) unit = [{ t: "lim", pos: op === LIM_LOW ? "low" : "upp", e: unit, lim: arg }];
      else unit = [{ t: op === "^" ? "sup" : "sub", base: unit, script: arg }];
    }
    push(unit);
  }
  return out;
}

/** Rewrite an m:oMath element's content from linear text. Formatting the
 * linear text does not spell out is carried over from what was there. */
export function setMathLinear(doc: DocxDocument, oMathEl: XmlElement, text: string): boolean {
  // Blank input is refused, but text that merely BEGINS or ends with a space
  // is written as given: equations really do start with one (" dt"), and
  // trimming here would make isLinearSafe promise a round-trip it then broke.
  if (!text.trim()) return false;
  const children = reemit(oMathEl, text);
  if (!children) return false;
  oMathEl.children = children;
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

/** The OMML element a node writes itself as. */
function tagOf(node: MathNode): string {
  switch (node.t) {
    case "run": return "r";
    case "sup": return "sSup";
    case "sub": return "sSub";
    case "frac": return "f";
    case "rad": return "rad";
    case "nary": return "nary";
    case "dlm": return "d";
    case "mat": return "m";
    case "eqarr": return "eqArr";
    case "acc": return "acc";
    case "grp": return "groupChr";
    case "lim": return node.pos === "low" ? "limLow" : "limUpp";
  }
}

/**
 * Put rebuilt content into one slot (an m:e, m:num, m:sup …), keeping the
 * property children Word hangs off the SLOT rather than off its parent — an
 * m:ctrlPr sitting inside an m:e — on whichever side of the content it was on.
 */
function fillSlot(target: XmlElement, nodes: MathNode[], m: string): void {
  const isProp = (c: XmlElement): boolean => localName(c.name).endsWith("Pr");
  const built = nodes.map((n) => buildOmml(n, m));
  const first = target.children.findIndex((c) => !isProp(c));
  if (first < 0) {
    // Nothing but properties (or nothing at all): the content follows them.
    target.children = [...target.children, ...built];
    return;
  }
  let last = target.children.length - 1;
  while (last > first && isProp(target.children[last])) last--;
  target.children = [...target.children.slice(0, first), ...built, ...target.children.slice(last + 1)];
}

/**
 * Re-emit a node into its own source element: the content slots take the
 * rebuilt children and every other child — m:fPr, m:naryPr, m:ctrlPr, a
 * matrix's m:mPr — stays exactly as Word wrote it. graft() only hands a node
 * a source whose spelled-out parts still match, so the properties kept here
 * always still describe the node.
 */
function intoTemplate(node: MathNode, src: XmlElement, m: string): XmlElement {
  const out = cloneXml(src);
  const slot = (name: string, nodes: MathNode[]): void => {
    const target = out.children.find((c) => localName(c.name) === name);
    if (target) fillSlot(target, nodes, m);
  };
  const parts = (name: string): XmlElement[] => out.children.filter((c) => localName(c.name) === name);
  switch (node.t) {
    case "run": {
      // The source run keeps its m:rPr, w:rPr and anything else it carries
      // (a w:br, say); only the text it spells moves.
      const texts = parts("t");
      if (texts.length === 0) {
        // A run that never held text (it carries a w:br) stays that way.
        if (node.text) out.children.push(el(`${m}t`, [], node.text));
      } else {
        texts[0].text = node.text;
        for (const extra of texts.slice(1)) out.children.splice(out.children.indexOf(extra), 1);
      }
      return out;
    }
    case "sup": slot("e", node.base); slot("sup", node.script); return out;
    case "sub": slot("e", node.base); slot("sub", node.script); return out;
    case "frac": slot("num", node.num); slot("den", node.den); return out;
    case "rad": slot("deg", node.deg ?? []); slot("e", node.e); return out;
    case "nary": slot("sub", node.sub); slot("sup", node.sup); slot("e", node.e); return out;
    case "acc":
    case "grp": slot("e", node.e); return out;
    case "lim": slot("e", node.e); slot("lim", node.lim); return out;
    case "dlm": {
      const es = parts("e");
      for (let k = 0; k < es.length && k < node.e.length; k++) fillSlot(es[k], node.e[k], m);
      return out;
    }
    case "eqarr": {
      const es = parts("e");
      for (let k = 0; k < es.length && k < node.rows.length; k++) fillSlot(es[k], node.rows[k], m);
      return out;
    }
    case "mat": {
      const rows = parts("mr");
      for (let r = 0; r < rows.length && r < node.rows.length; r++) {
        const cells = rows[r].children.filter((c) => localName(c.name) === "e");
        for (let c = 0; c < cells.length && c < node.rows[r].length; c++) {
          fillSlot(cells[c], node.rows[r][c], m);
        }
      }
      return out;
    }
  }
}

function buildOmml(node: MathNode, m: string): XmlElement {
  if (node.src && localName(node.src.name) === tagOf(node)) return intoTemplate(node, node.src, m);
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
 *
 * CROSS-PART moves are refused. Each editable root declares the `m` namespace
 * for itself (insertMathAt adds `xmlns:m` to the root it inserts into), so
 * carrying an equation from the body into a header would land OMML in a part
 * that never declared the namespace — invalid XML, and in a room every replica
 * would produce it identically. Same-part only, in both editing modes.
 */
export function moveMath(doc: DocxDocument, oMathEl: XmlElement, t: XmlElement, offset: number): boolean {
  const curParent = doc.findParentOf(oMathEl);
  const rEl = doc.findParentOf(t);
  const pEl = rEl && doc.findParentOf(rEl);
  if (!curParent || !rEl || !pEl || localName(rEl.name) !== "r") return false;
  if (rEl === oMathEl || curParent === oMathEl) return false;
  const holds = (root: XmlElement, needle: XmlElement): boolean =>
    root === needle || root.children.some((c) => holds(c, needle));
  const roots = doc.editableRoots();
  if (roots.find((r) => holds(r, oMathEl)) !== roots.find((r) => holds(r, t))) return false;
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

/**
 * Derive the math AST from a live oMath element (reflects current XML). Every
 * node keeps the element it came from, so an edit can be written back through
 * it (see buildOmml) without losing the formatting the linear text omits.
 *
 * OMML this does not model \u2014 m:func, m:sSubSup and the rest \u2014 falls through to
 * the recursion at the bottom and loses its wrapper. That is deliberate: the
 * lost wrapper makes the re-emitted equation differ from the original, which
 * is exactly what isLinearSafe looks for before it lets anyone edit.
 */
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
  // One node per m:r, never merged across runs: each run holds its own fonts,
  // and a run with no m:t at all (a w:br) still has to come back on save.
  if (ln === "r") {
    const text = e.children.filter((c) => localName(c.name) === "t").map((c) => c.text).join("");
    return [{ t: "run", text, src: e }];
  }
  if (ln === "f") return [{ t: "frac", num: kids("num"), den: kids("den"), src: e }];
  if (ln === "nary") {
    return [{ t: "nary", chr: chrAttr("naryPr", "chr", "\u222b"), sub: kids("sub"), sup: kids("sup"), e: kids("e"), src: e }];
  }
  if (ln === "d") {
    const pr = e.children.find((c) => localName(c.name) === "dPr");
    const chr = (name: string, dflt: string) => {
      const c = pr?.children.find((ch) => localName(ch.name) === name);
      const k = c && Object.keys(c.attrs).find((key) => localName(key) === "val");
      return c && k ? c.attrs[k] : dflt;
    };
    const parts = e.children.filter((c) => localName(c.name) === "e").map(ommlToNodes);
    return [{ t: "dlm", beg: chr("begChr", "("), end: chr("endChr", ")"), e: parts, src: e }];
  }
  if (ln === "m" && e.children.some((c) => localName(c.name) === "mr")) {
    const rows = e.children
      .filter((c) => localName(c.name) === "mr")
      .map((mr) => mr.children.filter((c) => localName(c.name) === "e").map(ommlToNodes));
    return [{ t: "mat", rows, src: e }];
  }
  if (ln === "acc") return [{ t: "acc", chr: chrAttr("accPr", "chr", "\u0302"), e: kids("e"), src: e }];
  if (ln === "groupChr") {
    return [
      {
        t: "grp",
        chr: chrAttr("groupChrPr", "chr", "\u23df"),
        pos: chrAttr("groupChrPr", "pos", "bot") === "top" ? "top" : "bot",
        vertJc: chrAttr("groupChrPr", "vertJc", "bot") === "top" ? "top" : "bot",
        e: kids("e"),
        src: e,
      },
    ];
  }
  if (ln === "limLow" || ln === "limUpp") {
    return [{ t: "lim", pos: ln === "limLow" ? "low" : "upp", e: kids("e"), lim: kids("lim"), src: e }];
  }
  if (ln === "eqArr") {
    const rows = e.children.filter((c) => localName(c.name) === "e").map(ommlToNodes);
    return [{ t: "eqarr", rows, src: e }];
  }
  if (ln === "sSup") return [{ t: "sup", base: kids("e"), script: kids("sup"), src: e }];
  if (ln === "sSub") return [{ t: "sub", base: kids("e"), script: kids("sub"), src: e }];
  if (ln === "rad") {
    const hide = chrAttr("radPr", "degHide", "0");
    const deg = hide === "1" || hide === "true" || hide === "on" ? [] : kids("deg");
    return deg.length ? [{ t: "rad", e: kids("e"), deg, src: e }] : [{ t: "rad", e: kids("e"), src: e }];
  }
  return e.children.flatMap(ommlToNodes);
}

/** The math AST currently in an oMath element (for prefilling the editor). */
export function mathLinearOf(doc: DocxDocument, oMathEl: XmlElement): string {
  void doc;
  return linearizeMath(ommlToNodes(oMathEl));
}

/**
 * Same kind and same spelled-out parts \u2014 the chr, the delimiters, the row and
 * column counts, whether the root shows an index. Two such nodes describe the
 * same construct, so the older one's source element still fits the newer one.
 */
function sameShape(a: MathNode, b: MathNode): boolean {
  if (a.t !== b.t) return false;
  switch (a.t) {
    case "run":
    case "sup":
    case "sub":
    case "frac":
      return true;
    case "rad":
      return !!a.deg?.length === !!(b as typeof a).deg?.length;
    case "nary":
      return a.chr === (b as typeof a).chr;
    case "dlm": {
      const y = b as typeof a;
      return a.beg === y.beg && a.end === y.end && a.e.length === y.e.length;
    }
    case "mat": {
      const y = b as typeof a;
      return a.rows.length === y.rows.length && a.rows.every((row, r) => row.length === y.rows[r].length);
    }
    case "eqarr":
      return a.rows.length === (b as typeof a).rows.length;
    case "acc":
    case "grp":
      return a.chr === (b as typeof a).chr;
    case "lim":
      return a.pos === (b as typeof a).pos;
  }
}

/** A stretch of consecutive runs, or one structural node. Runs group because
 * the reader keeps one node per m:r while the parser makes one node per
 * stretch of text; only the stretch as a whole lines up between them. */
type Span = { runs: MathNode[] } | { node: MathNode };

function spansOf(nodes: MathNode[]): Span[] {
  const out: Span[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (n.t === "run" && last && "runs" in last) last.runs.push(n);
    else out.push(n.t === "run" ? { runs: [n] } : { node: n });
  }
  return out;
}

const spanText = (runs: MathNode[]): string => runs.map((n) => (n.t === "run" ? n.text : "")).join("");

/**
 * Carry the source elements of `was` (read out of the document) onto `now`
 * (parsed back from the edited text), position by position. A node inherits a
 * source only while sameShape holds all the way down to it, so a restructured
 * subtree is built from scratch rather than inheriting properties that no
 * longer describe it. Anything that fails to line up is left as parsed.
 */
function graft(now: MathNode[], was: MathNode[]): MathNode[] {
  const fresh = spansOf(now);
  const old = spansOf(was);
  const out: MathNode[] = [];
  let f = 0;
  let o = 0;
  while (f < fresh.length || o < old.length) {
    const a = fresh[f];
    const b = old[o];
    // The reader keeps runs that hold no text (a lone w:br); the linear form
    // says nothing about them, so they simply come back where they were.
    if (b && "runs" in b && spanText(b.runs) === "" && (!a || !("runs" in a))) {
      out.push(...b.runs);
      o++;
      continue;
    }
    if (!a || !b || "runs" in a !== "runs" in b) return now;
    if ("runs" in a && "runs" in b) {
      // Untouched text goes back as the very runs it came from, so their fonts
      // and m:rPr survive byte for byte.
      if (spanText(a.runs) === spanText(b.runs)) out.push(...b.runs);
      else {
        // Edited text lands in one run, formatted like the first one that
        // carried text before.
        const template = b.runs.find((n) => n.t === "run" && n.text.length > 0) ?? b.runs[0];
        out.push(...a.runs.map((n, index) => (index === 0 ? { ...n, src: template.src } : n)));
      }
    } else if (!("runs" in a) && !("runs" in b)) {
      out.push(sameShape(a.node, b.node) ? withSource(a.node, b.node) : a.node);
    }
    f++;
    o++;
  }
  return out;
}

/** `now` with `was`'s source element and its slots grafted in turn. */
function withSource(now: MathNode, was: MathNode): MathNode {
  const src = was.src;
  switch (now.t) {
    case "run":
      return { ...now, src };
    case "sup":
    case "sub": {
      const w = was as typeof now;
      return { ...now, src, base: graft(now.base, w.base), script: graft(now.script, w.script) };
    }
    case "frac": {
      const w = was as typeof now;
      return { ...now, src, num: graft(now.num, w.num), den: graft(now.den, w.den) };
    }
    case "rad": {
      const w = was as typeof now;
      return { ...now, src, e: graft(now.e, w.e), ...(now.deg ? { deg: graft(now.deg, w.deg ?? []) } : {}) };
    }
    case "nary": {
      const w = was as typeof now;
      return { ...now, src, sub: graft(now.sub, w.sub), sup: graft(now.sup, w.sup), e: graft(now.e, w.e) };
    }
    case "dlm": {
      const w = was as typeof now;
      return { ...now, src, e: now.e.map((part, index) => graft(part, w.e[index])) };
    }
    case "mat": {
      const w = was as typeof now;
      return { ...now, src, rows: now.rows.map((row, r) => row.map((cell, c) => graft(cell, w.rows[r][c]))) };
    }
    case "eqarr": {
      const w = was as typeof now;
      return { ...now, src, rows: now.rows.map((row, r) => graft(row, w.rows[r])) };
    }
    case "acc":
    case "grp": {
      const w = was as typeof now;
      return { ...now, src, e: graft(now.e, w.e) };
    }
    case "lim": {
      const w = was as typeof now;
      return { ...now, src, e: graft(now.e, w.e), lim: graft(now.lim, w.lim) };
    }
  }
}

/** The namespace prefix an oMath element writes its children with. */
function mathPrefix(oMathEl: XmlElement): string {
  return oMathEl.name.includes(":") ? oMathEl.name.slice(0, oMathEl.name.indexOf(":") + 1) : "m:";
}

/** The OMML an edited linear string produces for this equation, or null when
 * the string says nothing: parse it, carry over the sources that still fit,
 * write it back out. */
function reemit(oMathEl: XmlElement, text: string): XmlElement[] | null {
  const nodes = parseMathLinear(text);
  if (nodes.length === 0) return null;
  const m = mathPrefix(oMathEl);
  return graft(nodes, ommlToNodes(oMathEl)).map((n) => buildOmml(n, m));
}

/**
 * True when the equation survives a trip out to linear text and back: the
 * OMML it re-emits is the OMML it started from, byte for byte. That is the
 * whole promise the editor makes \u2014 that opening an equation and saving it
 * unchanged changes nothing \u2014 so it is measured rather than reasoned about.
 * Equations built from OMML the linear form cannot name (m:func, m:sSubSup)
 * re-emit differently and open read-only.
 */
export function isLinearSafe(oMathEl: XmlElement): boolean {
  const text = linearizeMath(ommlToNodes(oMathEl));
  if (!text.trim()) return false;
  const children = reemit(oMathEl, text);
  return !!children && serializeXml({ ...oMathEl, children }) === serializeXml(oMathEl);
}
