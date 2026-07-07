import { DocxDocument } from "../docx.js";
import { MathNode } from "../model.js";
import { XmlElement, localName } from "../xml.js";

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
      case "rad":
        out += "√" + group(n.e);
        break;
    }
  }
  return out;
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

  /** One unit: a group, a root, or a single character (as a run). */
  const parseUnit = (): MathNode[] => {
    const ch = input[i];
    if (ch === "{") return parseGroup();
    if (ch === "√") {
      i++;
      const e = input[i] === "{" ? parseGroup() : parseUnit();
      return [{ t: "rad", e }];
    }
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
    case "rad":
      return el(`${m}rad`, [el(`${m}e`, node.e.map((n) => buildOmml(n, m)))]);
  }
}

/** The math AST currently in an oMath element (for prefilling the editor). */
export function mathLinearOf(doc: DocxDocument, oMathEl: XmlElement): string {
  void doc;
  // Re-derive from XML so the text reflects the current state.
  const parse = (e: XmlElement): MathNode[] => {
    const ln = localName(e.name);
    const kids = (name: string): MathNode[] => {
      const c = e.children.find((ch) => localName(ch.name) === name);
      return c ? parse(c) : [];
    };
    if (ln === "f") return [{ t: "frac", num: kids("num"), den: kids("den") }];
    if (ln === "sSup") return [{ t: "sup", base: kids("e"), script: kids("sup") }];
    if (ln === "sSub") return [{ t: "sub", base: kids("e"), script: kids("sub") }];
    if (ln === "rad") return [{ t: "rad", e: kids("e") }];
    if (ln === "t") return e.text ? [{ t: "run", text: e.text }] : [];
    const out: MathNode[] = [];
    for (const c of e.children) {
      for (const n of parse(c)) {
        const last = out[out.length - 1];
        if (n.t === "run" && last && last.t === "run") last.text += n.text;
        else out.push(n);
      }
    }
    return out;
  };
  return linearizeMath(parse(oMathEl));
}
