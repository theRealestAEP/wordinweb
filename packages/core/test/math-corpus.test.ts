import { describe, expect, it } from "vitest";
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { DocxDocument } from "../src/docx.js";
import { serializeXml, localName, type XmlElement } from "../src/xml.js";
import { isLinearSafe, mathLinearOf, setMathLinear } from "../src/edit/math.js";

/**
 * The equation editor's real gate, over Word's own OMML rather than OMML we
 * wrote ourselves: every equation the linear form claims it can edit must
 * re-emit byte for byte after a trip out to text and back.
 *
 * Opt-in, like the agent's corpus audit — the fixtures live in the parity
 * checkout beside this one.
 */
const enabled = process.env.WORDINWEB_FIXTURE_AUDIT === "1";

function corpusRoot(): string | null {
  for (const candidate of [
    process.env.WORDINWEB_FIXTURES,
    resolve(process.cwd(), "../../wordinweb-parity"),
    resolve(process.cwd(), "../../../wordinweb-parity"),
    resolve(process.cwd(), "../wordinweb-parity"),
  ]) {
    if (candidate && existsSync(resolve(candidate, "apps/demo/public/fixtures"))) {
      return resolve(candidate, "apps/demo/public/fixtures");
    }
  }
  return null;
}

/** The OMML-dense fixtures, and the totals measured when this landed. Raising
 * `editable` is welcome; dropping below it means equations that used to open
 * for editing no longer do. */
const FIXTURES = [
  // Dense refuses four m:func equations and one whose only content is a space;
  // chem refuses two m:sSubSup. Nothing else in the corpus is read-only.
  { file: "wild2-math-omml-dense.docx", equations: 38, editable: 33 },
  { file: "wild2-sci-chem-omml.docx", equations: 9, editable: 7 },
  { file: "parity-math.docx", equations: 1, editable: 1 },
  { file: "parity-math2.docx", equations: 4, editable: 4 },
  { file: "probe2-math-matrices.docx", equations: 7, editable: 7 },
];

function equationsIn(doc: DocxDocument): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "oMath") {
      out.push(el);
      return;
    }
    for (const child of el.children) walk(child);
  };
  for (const root of doc.editableRoots()) walk(root);
  return out;
}

describe.runIf(enabled)("math linear round-trip over the OMML corpus", () => {
  const root = corpusRoot();

  it("finds the parity fixtures", () => {
    expect(root, "Set WORDINWEB_FIXTURES to the wordinweb-parity checkout").toBeTruthy();
  });

  for (const fixture of FIXTURES) {
    it(`${fixture.file}: every editable equation re-emits byte for byte`, () => {
      const doc = DocxDocument.load(readFileSync(resolve(root!, fixture.file)));
      const equations = equationsIn(doc);
      expect(equations.length).toBe(fixture.equations);

      let editable = 0;
      for (const equation of equations) {
        if (!isLinearSafe(equation)) continue;
        editable++;
        const before = serializeXml(equation);
        expect(setMathLinear(doc, equation, mathLinearOf(doc, equation))).toBe(true);
        // The whole promise: opening an equation and saving it unchanged
        // changes nothing — not a font, not an m:ctrlPr, not a limit position.
        expect(serializeXml(equation), `${fixture.file}: ${before.slice(0, 200)}`).toBe(before);
      }
      expect(editable).toBeGreaterThanOrEqual(fixture.editable);
    });
  }
});
