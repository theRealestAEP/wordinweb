import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  evaluateTableFormula,
  formatFormulaNumber,
  formulaInstruction,
  insertTableFormula,
  isValidFormulaInstruction,
  parseTableFormula,
} from "../src/edit/formula.js";
import { isInsertableFieldInstruction } from "../src/edit/fields.js";
import { updateFields } from "../src/edit/update-fields.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** A table whose rows are cell-text arrays. */
function tbl(rows: string[][]): string {
  const tr = (cells: string[]) =>
    `<w:tr>${cells.map((c) => `<w:tc><w:p><w:r><w:t>${c}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`;
  return `<w:tbl><w:tblPr/><w:tblGrid/>${rows.map(tr).join("")}</w:tbl>`;
}

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** The w:p of the cell at rowIdx/cellIdx of the document's first table. */
function cellParagraph(doc: DocxDocument, rowIdx: number, cellIdx: number): XmlElement {
  const find = (el: XmlElement): XmlElement | null => {
    if (localName(el.name) === "tbl") return el;
    for (const c of el.children) {
      const found = find(c);
      if (found) return found;
    }
    return null;
  };
  const table = find(doc.docRoot)!;
  const tr = table.children.filter((c) => localName(c.name) === "tr")[rowIdx];
  const tc = tr.children.filter((c) => localName(c.name) === "tc")[cellIdx];
  return tc.children.find((c) => localName(c.name) === "p")!;
}

describe("formula parsing", () => {
  it("accepts the documented grammar", () => {
    for (const instr of [
      "=SUM(ABOVE)",
      "=sum(left)",
      "=AVERAGE(A1:B2)",
      "=COUNT(BELOW)",
      "=MAX(A1,B2,5)",
      "=MIN(RIGHT)",
      "=PRODUCT(A1:A3)",
      "=A1+B2",
      "=2*(3+4)^2",
      "=-A1/2",
      '=SUM(ABOVE) \\# "#,##0.00"',
      "=SUM(ABOVE) \\# 0.00",
    ]) {
      expect(isValidFormulaInstruction(instr), instr).toBe(true);
      expect(isInsertableFieldInstruction(instr), instr).toBe(true);
    }
  });

  it("refuses what the simple tier does not model", () => {
    for (const instr of [
      "=IF(A1>1,1,0)", // boolean functions
      "=ABS(-1)", // out-of-set function
      "=A1>B1", // comparisons
      "=DEFINED(x)",
      "=bookmarkName",
      "=", // empty
      "=SUM(ABOVE", // unbalanced
      "=SUM()", // missing argument
      `=SUM(ABOVE) \\# "${"#".repeat(40)}"`, // over-long picture
      "=SUM(ABOVE) \\* MERGEFORMAT", // other switches
      `=1 \\# "a\\"b"`, // quote smuggling
    ]) {
      expect(isValidFormulaInstruction(instr), instr).toBe(false);
      expect(isInsertableFieldInstruction(instr), instr).toBe(false);
    }
  });

  it("extracts the \\# picture", () => {
    const parsed = parseTableFormula('=SUM(ABOVE) \\# "#,##0.00"');
    expect(parsed?.numFmt).toBe("#,##0.00");
    expect(parseTableFormula("=SUM(ABOVE)")?.numFmt).toBeUndefined();
  });
});

describe("formula evaluation", () => {
  const grid = [
    ["10", "20", "x"],
    ["1.5", "2.5", "note"],
    ["", "", ""],
  ];

  const evalAt = (formula: string, row = 2, cell = 0): string | null => {
    const doc = load(tbl(grid) + p("after"));
    return evaluateTableFormula(doc, cellParagraph(doc, row, cell), formula);
  };

  it("sums a directional scan, stopping at the first empty cell", () => {
    // From row 2 col 0 upward: row 1 is "1.5", row 0 is "10".
    expect(evalAt("=SUM(ABOVE)")).toBe("11.5");
    // A gap stops the scan: from row 2 col 2 upward the neighbor ("") stops it.
    expect(evalAt("=SUM(ABOVE)", 2, 2)).toBe("0");
    // Text counts as 0 but continues the scan (row 1 col 2 is "note").
    expect(evalAt("=SUM(ABOVE)", 2, 2)).toBe("0");
    expect(evalAt("=SUM(LEFT)", 0, 2)).toBe("30");
  });

  it("evaluates cell references, ranges, and arithmetic", () => {
    expect(evalAt("=A1+B2")).toBe("12.5");
    expect(evalAt("=SUM(A1:B2)")).toBe("34");
    expect(evalAt("=AVERAGE(A1:B1)")).toBe("15");
    expect(evalAt("=COUNT(A1:C2)")).toBe("6");
    expect(evalAt("=MAX(A1:B2)")).toBe("20");
    expect(evalAt("=MIN(A1:B2)")).toBe("1.5");
    expect(evalAt("=PRODUCT(A1,B1)")).toBe("200");
    expect(evalAt("=2*(3+4)^2")).toBe("98");
    // An out-of-table reference reads as an empty cell: 0.
    expect(evalAt("=Z99")).toBe("0");
  });

  it("parses cell values locale-free, the sortTableRows rule", () => {
    const doc = load(tbl([["$1,234.50"], ["2"], [""]]) + p("after"));
    expect(evaluateTableFormula(doc, cellParagraph(doc, 2, 0), "=SUM(ABOVE)")).toBe("1236.5");
  });

  it("paints Word's zero-divide text", () => {
    expect(evalAt("=1/0")).toBe("!Zero Divide");
    expect(evalAt("=AVERAGE(A3:C3)")).toBe("!Zero Divide"); // empty range: no values
  });

  it("applies the \\# picture", () => {
    expect(formatFormulaNumber(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(formatFormulaNumber(1234.5, "0")).toBe("1235");
    expect(formatFormulaNumber(0.5, "0%")).toBe("50%");
    expect(formatFormulaNumber(-3, "$#,##0.00;($#,##0.00)")).toBe("($3.00)");
    expect(formatFormulaNumber(-3, "0.0")).toBe("-3.0");
    expect(formatFormulaNumber(0, "0;-0;zero")).toBe("zero");
    expect(formatFormulaNumber(3.14159, "0.##")).toBe("3.14");
    expect(evalAt('=SUM(A1:B1) \\# "#,##0.00"')).toBe("30.00");
  });
});

describe("insertTableFormula + updateFields", () => {
  it("writes a fldSimple with the evaluated cached result", () => {
    const doc = load(tbl([["10"], ["20"], [""]]) + p("after"));
    const target = cellParagraph(doc, 2, 0);
    expect(insertTableFormula(doc, target, "SUM(ABOVE)", "#,##0.00")).toBe(true);
    const xml = serializeXml(target);
    expect(xml).toContain('w:instr=" =SUM(ABOVE) \\# &quot;#,##0.00&quot; "');
    expect(xml).toContain(">30.00<");
  });

  it("refuses outside a table and refuses bad formulas", () => {
    const doc = load(p("no table here"));
    const para = doc.docRoot.children
      .find((c) => localName(c.name) === "body")!
      .children.find((c) => localName(c.name) === "p")!;
    expect(insertTableFormula(doc, para, "SUM(ABOVE)")).toBe(false);
    const doc2 = load(tbl([["1"], [""]]) + p("x"));
    expect(insertTableFormula(doc2, cellParagraph(doc2, 1, 0), "IF(1,2,3)")).toBe(false);
  });

  it("recomputes the cached result on updateFields after cells change", () => {
    const doc = load(tbl([["10"], ["20"], [""]]) + p("after"));
    insertTableFormula(doc, cellParagraph(doc, 2, 0), "SUM(ABOVE)");
    expect(serializeXml(doc.docRoot)).toContain(">30<");
    // Change a contributing cell's text, then run the update pass.
    const cell = cellParagraph(doc, 1, 0);
    const t = cell.children.find((c) => localName(c.name) === "r")!.children.find((c) => localName(c.name) === "t")!;
    t.text = "25";
    doc.refresh();
    expect(updateFields(doc)).toBe(true);
    expect(serializeXml(doc.docRoot)).toContain(">35<");
  });

  it("builds the exact instruction the validator accepts", () => {
    expect(formulaInstruction("SUM(ABOVE)")).toBe("=SUM(ABOVE)");
    expect(formulaInstruction("=A1+1", "0.00")).toBe('=A1+1 \\# "0.00"');
    expect(formulaInstruction("IF(1,2,3)")).toBe(null);
    expect(formulaInstruction("SUM(ABOVE)", 'bad"fmt')).toBe(null);
  });
});
