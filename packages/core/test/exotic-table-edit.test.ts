import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { applyTableOp, cellContextOf } from "../src/edit/tables.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { XmlElement, child, localName } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** A 2x2 floating table (w:tblpPr) with the given tblPr extras. */
function floatingTable(): string {
  return `<w:tbl>
    <w:tblPr>
      <w:tblpPr w:leftFromText="180" w:rightFromText="180" w:topFromText="180" w:bottomFromText="180"
        w:vertAnchor="page" w:horzAnchor="page" w:tblpX="1500" w:tblpY="1500"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:color="000000"/>
        <w:left w:val="single" w:sz="6" w:color="000000"/><w:right w:val="single" w:sz="6" w:color="000000"/>
        <w:insideH w:val="single" w:sz="6" w:color="000000"/><w:insideV w:val="single" w:sz="6" w:color="000000"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("r1c1")}</w:tc>
          <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("r1c2")}</w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("r2c1")}</w:tc>
          <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("r2c2")}</w:tc></w:tr>
  </w:tbl>`;
}

/** A 2x2 old-style separated-border table (w:tblCellSpacing). */
function cellSpacingTable(): string {
  return `<w:tbl>
    <w:tblPr>
      <w:tblCellSpacing w:w="60" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:color="000000"/>
        <w:left w:val="single" w:sz="6" w:color="000000"/><w:right w:val="single" w:sz="6" w:color="000000"/>
        <w:insideH w:val="single" w:sz="6" w:color="000000"/><w:insideV w:val="single" w:sz="6" w:color="000000"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("a1")}</w:tc>
          <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("a2")}</w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("b1")}</w:tc>
          <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>${p("b2")}</w:tc></w:tr>
  </w:tbl>`;
}

function findT(root: XmlElement, needle: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (e: XmlElement) => {
    if (found) return;
    if (localName(e.name) === "t" && (e.text ?? "").includes(needle)) found = e;
    else e.children.forEach(walk);
  };
  walk(root);
  if (!found) throw new Error(`no w:t containing ${needle}`);
  return found;
}
function docRootOf(doc: DocxDocument): XmlElement {
  return (doc as unknown as { docRoot: XmlElement }).docRoot;
}
function tblOf(doc: DocxDocument, t: XmlElement): XmlElement {
  let cur: XmlElement | undefined = doc.findParentOf(t);
  while (cur) { if (localName(cur.name) === "tbl") return cur; cur = doc.findParentOf(cur); }
  throw new Error("no enclosing tbl");
}
function has(el: XmlElement, name: string): boolean {
  const walk = (e: XmlElement): boolean => localName(e.name) === name || e.children.some(walk);
  return walk(el);
}
function rows(tbl: XmlElement) { return tbl.children.filter((c) => localName(c.name) === "tr"); }
function cellsIn(tr: XmlElement) { return tr.children.filter((c) => localName(c.name) === "tc"); }
function gridCols(tbl: XmlElement) {
  const g = child(tbl, "tblGrid");
  return g ? g.children.filter((c) => localName(c.name) === "gridCol").length : 0;
}
/** Ordinal position of tblpPr / tblCellSpacing inside tblPr must stay first
 * (schema: tblpPr and tblCellSpacing precede tblBorders in EG_TblPrBase). */
function tblPrOrderOk(tbl: XmlElement): boolean {
  const pr = child(tbl, "tblPr");
  if (!pr) return false;
  const names = pr.children.map((c) => localName(c.name));
  const bordersIdx = names.indexOf("tblBorders");
  const floatIdx = names.indexOf("tblpPr");
  const spaceIdx = names.indexOf("tblCellSpacing");
  if (floatIdx !== -1 && !(floatIdx < bordersIdx)) return false;
  if (spaceIdx !== -1 && !(spaceIdx < bordersIdx)) return false;
  return true;
}

describe("editing exotic tables (floating tblpPr + old-style tblCellSpacing)", () => {
  it("resolves a cell context inside a floating table", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(floatingTable()) }));
    const t = findT(docRootOf(doc), "r1c1");
    const ctx = cellContextOf(doc, t);
    expect(ctx).not.toBeNull();
    expect(ctx!.rowIdx).toBe(0);
    expect(ctx!.cellIdx).toBe(0);
  });

  it("rowBelow / colRight / deleteCol keep the floating anchor and a consistent grid", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(floatingTable()) }));
    const t = findT(docRootOf(doc), "r1c1");
    const tbl = tblOf(doc, t);

    expect(applyTableOp(doc, t, "rowBelow")).toBe(true);
    expect(rows(tbl).length).toBe(3);
    expect(has(tbl, "tblpPr")).toBe(true);

    expect(applyTableOp(doc, t, "colRight")).toBe(true);
    expect(gridCols(tbl)).toBe(3);
    expect(rows(tbl).every((r) => cellsIn(r).length === 3)).toBe(true);
    expect(has(tbl, "tblpPr")).toBe(true);

    expect(applyTableOp(doc, t, "deleteCol")).toBe(true);
    expect(gridCols(tbl)).toBe(2);
    expect(rows(tbl).every((r) => cellsIn(r).length === 2)).toBe(true);
    expect(has(tbl, "tblpPr")).toBe(true);
    expect(tblPrOrderOk(tbl)).toBe(true);
  });

  it("rowBelow / colRight keep tblCellSpacing on the old-style table", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(cellSpacingTable()) }));
    const t = findT(docRootOf(doc), "a1");
    const tbl = tblOf(doc, t);

    expect(applyTableOp(doc, t, "rowBelow")).toBe(true);
    expect(rows(tbl).length).toBe(3);
    expect(has(tbl, "tblCellSpacing")).toBe(true);

    expect(applyTableOp(doc, t, "colRight")).toBe(true);
    expect(gridCols(tbl)).toBe(3);
    expect(rows(tbl).every((r) => cellsIn(r).length === 3)).toBe(true);
    expect(has(tbl, "tblCellSpacing")).toBe(true);
    expect(tblPrOrderOk(tbl)).toBe(true);
  });

  it("emits column + row resize grips for a floating table", () => {
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(p("lead") + floatingTable() + section) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const grips = result.pages.flatMap((pg) => pg.items).filter((i) => i.kind === "grip");
    const col = grips.filter((g) => g.kind === "grip" && g.axis === "col");
    const row = grips.filter((g) => g.kind === "grip" && g.axis === "row");
    // 2 columns -> 2 vertical boundaries; 2 rows -> 2 row boundaries.
    expect(col.length).toBe(2);
    expect(row.length).toBe(2);
  });
});
