import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  listTableStyles,
  readTableProperties,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  setTableLayoutMode,
  setTableLook,
  setTableStyle,
  setTableWidth,
  tableLookOf,
} from "../src/edit/tables.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { LayoutResult } from "../src/layout/types.js";
import { XmlElement, attr, child, localName, serializeXml } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/** A table style whose conditional formats are all distinguishable by colour,
 * so a rendered page says WHICH conditional format was applied and not merely
 * that one was. */
const STYLES_XML = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="GridBlue">
    <w:name w:val="Grid Blue"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="4472C4"/>
        <w:bottom w:val="single" w:sz="4" w:color="4472C4"/>
        <w:left w:val="single" w:sz="4" w:color="4472C4"/>
        <w:right w:val="single" w:sz="4" w:color="4472C4"/>
        <w:insideH w:val="single" w:sz="4" w:color="4472C4"/>
        <w:insideV w:val="single" w:sz="4" w:color="4472C4"/>
      </w:tblBorders>
      <w:tblStyleRowBandSize w:val="1"/>
      <w:tblStyleColBandSize w:val="1"/>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
      <w:rPr><w:b/></w:rPr>
      <w:tcPr><w:shd w:val="clear" w:fill="2F5496"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastRow">
      <w:tcPr><w:shd w:val="clear" w:fill="C00000"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="firstCol">
      <w:tcPr><w:shd w:val="clear" w:fill="00B050"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastCol">
      <w:tcPr><w:shd w:val="clear" w:fill="7030A0"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Horz">
      <w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Vert">
      <w:tcPr><w:shd w:val="clear" w:fill="FFF2CC"/></w:tcPr>
    </w:tblStylePr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

const CONTENT_TYPES_WITH_STYLES = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const DOC_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

/** rows x cols of text cells, each cell tagged r{row}c{col}. */
function tableXml(rows: number, cols: number, tblPrExtra = "", cellW = 2000): string {
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${cellW}"/>`).join("");
  const body = Array.from({ length: rows }, (_, r) => {
    const cells = Array.from(
      { length: cols },
      (_, c) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${cellW}"/></w:tcPr>${p(`r${r}c${c}`)}</w:tc>`,
    ).join("");
    return `<w:tr>${cells}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr>${tblPrExtra}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function load(bodyXml: string, withStyles = false): DocxDocument {
  const parts: Record<string, string> = { "word/document.xml": wrapDocument(bodyXml) };
  if (withStyles) {
    parts["word/styles.xml"] = STYLES_XML;
    parts["[Content_Types].xml"] = CONTENT_TYPES_WITH_STYLES;
    parts["word/_rels/document.xml.rels"] = DOC_RELS;
  }
  return DocxDocument.load(makeDocx(parts));
}

function layout(doc: DocxDocument): LayoutResult {
  return layoutDocument(doc, { measurer: new ApproxMeasurer() });
}

function findT(root: XmlElement, needle: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (e: XmlElement): void => {
    if (found) return;
    if (localName(e.name) === "t" && (e.text ?? "").includes(needle)) found = e;
    else e.children.forEach(walk);
  };
  walk(root);
  if (!found) throw new Error(`no w:t containing ${needle}`);
  return found;
}

function docRoot(doc: DocxDocument): XmlElement {
  return (doc as unknown as { docRoot: XmlElement }).docRoot;
}

function tblOf(doc: DocxDocument): XmlElement {
  let found: XmlElement | null = null;
  const walk = (e: XmlElement): void => {
    if (found) return;
    if (localName(e.name) === "tbl") found = e;
    else e.children.forEach(walk);
  };
  walk(docRoot(doc));
  if (!found) throw new Error("no w:tbl");
  return found;
}

/** The cell fills a page paints, in paint order. */
function fills(result: LayoutResult): string[] {
  return result.pages
    .flatMap((page) => page.items)
    .filter((item) => item.kind === "rect" && item.role === "table-fill")
    .map((item) => (item.kind === "rect" ? item.fill.toUpperCase() : ""));
}

/** Every table rule a page paints, as a comparable tuple. */
function rules(result: LayoutResult): string[] {
  return result.pages
    .flatMap((page) => page.items)
    .filter((item) => item.kind === "edge" && item.role === "table-rule")
    .map((item) =>
      item.kind === "edge"
        ? [
            item.x1.toFixed(2),
            item.y1.toFixed(2),
            item.x2.toFixed(2),
            item.y2.toFixed(2),
            item.border.style,
            item.border.color,
            item.border.width.toFixed(3),
          ].join("|")
        : "",
    )
    .sort();
}

/** Rendered column widths of the first table, from its resize grips. */
function renderedWidths(result: LayoutResult): number[] {
  for (const page of result.pages) {
    for (const item of page.items) {
      if (item.kind === "grip" && item.axis === "col" && item.renderedWidths) {
        return item.renderedWidths;
      }
    }
  }
  throw new Error("no column grips");
}

function tcPrOf(doc: DocxDocument, cellText: string): XmlElement | undefined {
  let tc: XmlElement | undefined = doc.findParentOf(findT(docRoot(doc), cellText));
  while (tc && localName(tc.name) !== "tc") tc = doc.findParentOf(tc);
  return tc ? child(tc, "tcPr") : undefined;
}

function xml(el: XmlElement | undefined): string {
  return el ? serializeXml(el) : "";
}

// ---------------------------------------------------------------------------
// Per-edge borders
// ---------------------------------------------------------------------------

describe("per-edge cell and table borders", () => {
  it("writes w:tcBorders for the requested edges only, in schema order", () => {
    const doc = load(tableXml(2, 2));
    const caret = findT(docRoot(doc), "r0c0");
    expect(
      setTableBorders(doc, caret, "cell", ["bottom", "top", "left"], {
        style: "double",
        sz: 12,
        color: "#FF0000",
        space: 2,
      }),
    ).toBe(true);
    const borders = child(tcPrOf(doc, "r0c0")!, "tcBorders")!;
    expect(borders.children.map((c) => localName(c.name))).toEqual(["top", "left", "bottom"]);
    expect(xml(child(borders, "top"))).toBe(
      `<w:top w:val="double" w:sz="12" w:space="2" w:color="FF0000"/>`,
    );
    // Only the addressed cell is touched.
    expect(child(tcPrOf(doc, "r0c1")!, "tcBorders")).toBeUndefined();
  });

  it("writes w:tblBorders at table scope, including the inside rules", () => {
    const doc = load(tableXml(2, 2));
    expect(
      setTableBorders(doc, findT(docRoot(doc), "r0c0"), "table", ["insideH", "insideV"], {
        style: "dashed",
        sz: 4,
      }),
    ).toBe(true);
    const borders = child(child(tblOf(doc), "tblPr")!, "tblBorders")!;
    expect(borders.children.map((c) => localName(c.name))).toEqual(["insideH", "insideV"]);
    expect(attr(child(borders, "insideH"), "color")).toBe("auto");
  });

  it("keeps tblBorders after tblInd and before tblLayout in tblPr", () => {
    const doc = load(
      tableXml(2, 2, `<w:tblInd w:w="120" w:type="dxa"/><w:tblLayout w:type="fixed"/>`),
    );
    setTableBorders(doc, findT(docRoot(doc), "r0c0"), "table", ["top"], { style: "single" });
    expect(child(tblOf(doc), "tblPr")!.children.map((c) => localName(c.name))).toEqual([
      "tblInd",
      "tblBorders",
      "tblLayout",
    ]);
  });

  it("distinguishes 'no border' from 'no direct border'", () => {
    // The table declares every rule, so an untouched cell inherits them.
    const tblPr = `<w:tblBorders>
      <w:top w:val="single" w:sz="8" w:color="000000"/>
      <w:bottom w:val="single" w:sz="8" w:color="000000"/>
      <w:left w:val="single" w:sz="8" w:color="000000"/>
      <w:right w:val="single" w:sz="8" w:color="000000"/>
      <w:insideH w:val="single" w:sz="8" w:color="000000"/>
      <w:insideV w:val="single" w:sz="8" w:color="000000"/>
    </w:tblBorders>`;
    const doc = load(tableXml(2, 2, tblPr));
    const inherited = rules(layout(doc)).length;

    // style "none" SUPPRESSES the inherited rule: w:val="nil", one fewer rule.
    setTableBorders(doc, findT(docRoot(doc), "r0c0"), "cell", ["top"], { style: "none" });
    expect(xml(child(child(tcPrOf(doc, "r0c0")!, "tcBorders")!, "top"))).toBe(`<w:top w:val="nil"/>`);
    const suppressed = rules(layout(doc)).length;
    expect(suppressed).toBe(inherited - 1);

    // A null spec REMOVES the edge, so the table's rule comes back.
    expect(setTableBorders(doc, findT(docRoot(doc), "r0c0"), "cell", ["top"], null)).toBe(true);
    expect(child(tcPrOf(doc, "r0c0")!, "tcBorders")).toBeUndefined();
    expect(rules(layout(doc)).length).toBe(inherited);
  });

  it("retires a w:start twin when it sets the left edge", () => {
    const doc = load(tableXml(1, 1));
    const tcPr = tcPrOf(doc, "r0c0")!;
    tcPr.children.push({
      name: "w:tcBorders",
      attrs: {},
      text: "",
      children: [
        { name: "w:start", attrs: { "w:val": "single", "w:sz": "24" }, text: "", children: [] },
      ],
    });
    setTableBorders(doc, findT(docRoot(doc), "r0c0"), "cell", ["left"], { style: "single", sz: 4 });
    const borders = child(tcPrOf(doc, "r0c0")!, "tcBorders")!;
    expect(borders.children.map((c) => localName(c.name))).toEqual(["left"]);
  });

  it("paints a cell diagonal as a path", () => {
    const doc = load(tableXml(1, 1));
    const before = layout(doc).pages.flatMap((pg) => pg.items).filter((i) => i.kind === "path").length;
    setTableBorders(doc, findT(docRoot(doc), "r0c0"), "cell", ["tl2br"], { style: "single", sz: 8 });
    const after = layout(doc).pages.flatMap((pg) => pg.items).filter((i) => i.kind === "path").length;
    expect(after).toBe(before + 1);
  });

  it("refuses edges the scope cannot carry", () => {
    const doc = load(tableXml(1, 1));
    // insideH is meaningless on a single cell and the renderer never reads it,
    // so writing it would be XML that does nothing.
    expect(
      setTableBorders(doc, findT(docRoot(doc), "r0c0"), "cell", ["insideH"], { style: "single" }),
    ).toBe(false);
    expect(
      setTableBorders(doc, findT(docRoot(doc), "r0c0"), "table", ["tl2br"], { style: "single" }),
    ).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Table style application
// ---------------------------------------------------------------------------

describe("table style application", () => {
  it("lists the table styles styles.xml defines", () => {
    expect(listTableStyles(load(tableXml(1, 1), true))).toEqual([{ id: "GridBlue", name: "Grid Blue" }]);
  });

  it("refuses a style the document does not define", () => {
    const doc = load(tableXml(1, 1), true);
    expect(setTableStyle(doc, tblOf(doc), "NoSuchStyle")).toBe(false);
    expect(setTableStyle(doc, tblOf(doc), "Normal")).toBe(false); // a PARAGRAPH style
    expect(child(child(tblOf(doc), "tblPr")!, "tblStyle")).toBeUndefined();
  });

  // The renderer's conditional-format resolution is proved against Word by the
  // parity corpus. What this pins is that APPLYING a style through the editing
  // surface lands the file in the state the corpus already covers: the same
  // document authored with w:tblStyle + w:tblLook in the XML must lay out
  // identically, item for item.
  it("renders exactly like the same table authored with the style in XML", () => {
    const authored = load(
      tableXml(
        4,
        4,
        `<w:tblStyle w:val="GridBlue"/>` +
          `<w:tblLook w:val="06A0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`,
      ),
      true,
    );
    const applied = load(tableXml(4, 4), true);
    expect(setTableStyle(applied, tblOf(applied), "GridBlue")).toBe(true);
    expect(
      setTableLook(applied, tblOf(applied), {
        firstRow: true,
        lastRow: true,
        firstColumn: true,
        lastColumn: false,
        bandedRows: true,
        bandedCols: false,
      }),
    ).toBe(true);

    const expectedFills = fills(layout(authored));
    // The fixture is only meaningful if the style actually paints: firstRow,
    // lastRow, firstCol and the horizontal band must all be visible.
    expect(new Set(expectedFills)).toEqual(
      new Set(["#2F5496", "#C00000", "#00B050", "#D9E2F3"]),
    );
    expect(fills(layout(applied))).toEqual(expectedFills);
    expect(rules(layout(applied))).toEqual(rules(layout(authored)));
  });

  it("writes every tblLook attribute, because a missing one reads as off", () => {
    const doc = load(tableXml(2, 2), true);
    setTableLook(doc, tblOf(doc), { bandedCols: false });
    const look = child(child(tblOf(doc), "tblPr")!, "tblLook")!;
    // Defaults are firstRow + firstColumn + both bands; only bandedCols moved.
    expect(xml(look)).toBe(
      `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`,
    );
    expect(tableLookOf(tblOf(doc))).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      bandedRows: true,
      bandedCols: false,
    });
  });

  it("round-trips the toggles through the legacy hex form", () => {
    const doc = load(tableXml(2, 2), true);
    // Only the hex bitmask, no attributes — the form older writers emit.
    const tblPr = child(tblOf(doc), "tblPr")!;
    tblPr.children.push({ name: "w:tblLook", attrs: { "w:val": "0460" }, text: "", children: [] });
    expect(tableLookOf(tblOf(doc))).toEqual({
      firstRow: true,
      lastRow: true,
      firstColumn: false,
      lastColumn: false,
      bandedRows: true,
      bandedCols: false,
    });
  });

  it("removes the style reference and its conditional paint", () => {
    const doc = load(tableXml(3, 3, `<w:tblStyle w:val="GridBlue"/>`), true);
    expect(fills(layout(doc)).length).toBeGreaterThan(0);
    expect(setTableStyle(doc, tblOf(doc), null)).toBe(true);
    expect(fills(layout(doc))).toEqual([]);
    expect(setTableStyle(doc, tblOf(doc), null)).toBe(false); // already gone
  });
});

// ---------------------------------------------------------------------------
// Numeric widths and the autofit interaction
// ---------------------------------------------------------------------------

describe("numeric table widths", () => {
  it("writes tblW in the unit asked for", () => {
    const doc = load(tableXml(1, 2));
    setTableWidth(doc, tblOf(doc), "pt", 216);
    expect(xml(child(child(tblOf(doc), "tblPr")!, "tblW"))).toBe(`<w:tblW w:w="4320" w:type="dxa"/>`);
    // Percent is fiftieths of a percent, the form Word emits.
    setTableWidth(doc, tblOf(doc), "pct", 80);
    expect(xml(child(child(tblOf(doc), "tblPr")!, "tblW"))).toBe(`<w:tblW w:w="4000" w:type="pct"/>`);
    setTableWidth(doc, tblOf(doc), "auto");
    expect(xml(child(child(tblOf(doc), "tblPr")!, "tblW"))).toBe(`<w:tblW w:w="0" w:type="auto"/>`);
  });

  it("sets one column exactly and re-totals the table", () => {
    const doc = load(tableXml(2, 3, `<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>`));
    expect(setTableColumnWidth(doc, tblOf(doc), 1, 180)).toBe(true); // 180pt = 3600tw
    const cols = child(tblOf(doc), "tblGrid")!.children.map((c) => attr(c, "w"));
    expect(cols).toEqual(["2000", "3600", "2000"]);
    // Every cell is re-stamped, so layout trusts the grid it is handed...
    expect(attr(child(tcPrOf(doc, "r0c1")!, "tcW"), "w")).toBe("3600");
    expect(attr(child(tcPrOf(doc, "r1c1")!, "tcW"), "w")).toBe("3600");
    // ...and the declared total follows the columns rather than contradicting them.
    expect(attr(child(child(tblOf(doc), "tblPr")!, "tblW"), "w")).toBe("7600");
    expect(setTableColumnWidth(doc, tblOf(doc), 9, 100)).toBe(false);
  });

  it("does not convert a percentage width into a fixed one", () => {
    const doc = load(tableXml(1, 2, `<w:tblW w:w="2500" w:type="pct"/>`));
    setTableColumnWidth(doc, tblOf(doc), 0, 100);
    expect(xml(child(child(tblOf(doc), "tblPr")!, "tblW"))).toBe(`<w:tblW w:w="2500" w:type="pct"/>`);
  });
});

describe("autofit and fixed layout interact", () => {
  /** An autofit table whose grid does NOT describe its content: the cells
   * declare no width, so layout measures them and the columns come out very
   * different from the even grid. That gap is what makes the freeze
   * observable. */
  function autofitDoc(): DocxDocument {
    const grid = `<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>`;
    const cell = (text: string) => `<w:tc>${p(text)}</w:tc>`;
    return load(
      `<w:tbl><w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>` +
        `<w:tr>${cell("r0c0 a very much longer first column indeed")}${cell("r0c1 short")}</w:tr>` +
        `<w:tr>${cell("r1c0 also long here")}${cell("r1c1 x")}</w:tr>` +
        `</w:tbl>`,
    );
  }

  it("freezes the RENDERED columns when it switches autofit to fixed", () => {
    const doc = autofitDoc();
    const before = renderedWidths(layout(doc));
    // The premise: the autofit columns are nothing like the even authored grid.
    expect(Math.abs(before[0] - before[1])).toBeGreaterThan(20);

    expect(setTableLayoutMode(doc, tblOf(doc), "fixed", before)).toBe(true);
    const after = renderedWidths(layout(doc));
    after.forEach((w, i) => expect(w).toBeCloseTo(before[i], 1));

    // The freeze is written down, not merely re-derived: grid, per-cell tcW
    // and the declared total all agree, which is what makes layout trust it.
    expect(attr(child(child(tblOf(doc), "tblPr")!, "tblLayout"), "type")).toBe("fixed");
    const gridTwips = child(tblOf(doc), "tblGrid")!.children.map((c) => Number(attr(c, "w")));
    expect(attr(child(child(tblOf(doc), "tblPr")!, "tblW"), "w")).toBe(
      String(gridTwips[0] + gridTwips[1]),
    );
    expect(attr(child(tcPrOf(doc, "r0c0")!, "tcW"), "w")).toBe(String(gridTwips[0]));
    expect(attr(child(tcPrOf(doc, "r1c1")!, "tcW"), "w")).toBe(String(gridTwips[1]));
  });

  it("survives a round trip through fixed and back to autofit", () => {
    const doc = autofitDoc();
    const original = renderedWidths(layout(doc));
    setTableLayoutMode(doc, tblOf(doc), "fixed", original);
    setTableLayoutMode(doc, tblOf(doc), "autofit");
    const back = renderedWidths(layout(doc));
    back.forEach((w, i) => expect(w).toBeCloseTo(original[i], 1));
  });

  it("re-measures from content when it switches fixed to autofit", () => {
    // A FIXED table with an even grid renders even columns whatever the text
    // says. Autofit must throw that away and measure the content instead.
    const grid = `<w:gridCol w:w="2500"/><w:gridCol w:w="2500"/>`;
    const cell = (text: string) =>
      `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2500"/></w:tcPr>${p(text)}</w:tc>`;
    const doc = load(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
        `<w:tblGrid>${grid}</w:tblGrid>` +
        `<w:tr>${cell("r0c0 a very much longer first column indeed")}${cell("r0c1 x")}</w:tr>` +
        `</w:tbl>`,
    );
    const fixed = renderedWidths(layout(doc));
    expect(fixed[0]).toBeCloseTo(fixed[1], 1);

    expect(setTableLayoutMode(doc, tblOf(doc), "autofit")).toBe(true);
    const auto = renderedWidths(layout(doc));
    expect(auto[0]).toBeGreaterThan(auto[1] + 20);

    // The three things that would have kept the frozen widths are all gone.
    expect(attr(child(child(tblOf(doc), "tblPr")!, "tblLayout"), "type")).toBe("autofit");
    expect(xml(child(child(tblOf(doc), "tblPr")!, "tblW"))).toBe(`<w:tblW w:w="0" w:type="auto"/>`);
    expect(child(tcPrOf(doc, "r0c0")!, "tcW")).toBeUndefined();
  });

  it("commits the existing grid when no rendered widths are supplied", () => {
    const doc = load(tableXml(1, 2, `<w:tblLayout w:type="autofit"/>`, 1500));
    setTableLayoutMode(doc, tblOf(doc), "fixed");
    expect(attr(child(child(tblOf(doc), "tblPr")!, "tblW"), "w")).toBe("3000");
    expect(attr(child(tcPrOf(doc, "r0c0")!, "tcW"), "w")).toBe("1500");
  });
});

// ---------------------------------------------------------------------------
// Percentage width against the cell margins
// ---------------------------------------------------------------------------

describe("percentage table width", () => {
  // A4 with 1in margins: an 11906tw page less 2880tw of margin leaves a
  // 9026tw text column, 601.73px at the engine's 15 twips per pixel.
  const CONTENT_WIDTH = (11906 - 2 * 1440) / 15;
  const SECTION =
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  /** 10pt left and right cell margins, the pair Word's default table uses. */
  const CELL_MAR =
    `<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar>`;
  /** The 26.67px the two 10pt margins add on each side of the table box. */
  const EDGE_MARGINS = (2 * 200) / 15;

  /** `mode` undefined leaves settings.xml out altogether, the shape that makes
   * Word fall back to its legacy table metrics. */
  function loadWithCompat(bodyXml: string, mode?: number): DocxDocument {
    const parts: Record<string, string> = { "word/document.xml": wrapDocument(bodyXml) };
    if (mode !== undefined) {
      parts["word/settings.xml"] =
        `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:compat><w:compatSetting w:name="compatibilityMode" w:val="${mode}"/></w:compat></w:settings>`;
    }
    return DocxDocument.load(makeDocx(parts));
  }

  const totalWidth = (doc: DocxDocument): number =>
    renderedWidths(layout(doc)).reduce((a, b) => a + b, 0);

  const pctTable = (layoutMode: "autofit" | "fixed"): string =>
    tableXml(
      1,
      2,
      `<w:tblW w:w="4500" w:type="pct"/><w:tblLayout w:type="${layoutMode}"/>` + CELL_MAR,
      3000,
    ) + SECTION;

  it("fits the cell margins inside a fixed 90% table under compat 15", () => {
    // Word 2013 metrics: the box is 0.90 x the text column and the margins
    // inset the cell text within it.
    const doc = loadWithCompat(pctTable("fixed"), 15);
    expect(totalWidth(doc)).toBeCloseTo(0.9 * CONTENT_WIDTH, 2);
  });

  for (const [name, mode] of [
    ["compat 12", 12],
    ["no declared compat", undefined],
  ] as const) {
    it(`adds the cell margins around a fixed 90% table under ${name}`, () => {
      // Legacy metrics: the box starts a cell margin left of the column, so
      // the percentage resolves against the column plus both margins.
      const doc = loadWithCompat(pctTable("fixed"), mode);
      expect(totalWidth(doc)).toBeCloseTo(0.9 * (CONTENT_WIDTH + EDGE_MARGINS), 2);
    });
  }

  it("leaves an autofit-layout percentage table on the bare text column", () => {
    // The allowance belongs to fixed layout; autofit resolves against the
    // column in either compatibility mode.
    for (const mode of [15, 12, undefined]) {
      const doc = loadWithCompat(pctTable("autofit"), mode);
      expect(totalWidth(doc)).toBeCloseTo(0.9 * CONTENT_WIDTH, 2);
    }
  });

  it("keeps a dxa width at its declared size whatever the margins are", () => {
    const doc = load(tableXml(1, 2, `<w:tblW w:w="6000" w:type="dxa"/>`, 3000) + SECTION);
    const before = totalWidth(doc);
    expect(before).toBeCloseTo(6000 / 15, 2);
    setTableCellMargins(doc, findT(docRoot(doc), "r0c0"), "table", { left: 10, right: 10 });
    expect(totalWidth(doc)).toBeCloseTo(before, 2);
  });
});

// ---------------------------------------------------------------------------
// How a percentage width is split between the columns
// ---------------------------------------------------------------------------

describe("percentage column distribution", () => {
  // Letter with 1in margins: a 9360tw text column, 624px at 15 twips per pixel.
  const CONTENT_WIDTH = (12240 - 2 * 1440) / 15;
  const SECTION =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

  /** A compat-15 percentage table over an explicit grid. `marginTw` is one
   * side's cell margin, so a column reserves twice it. */
  function pctGrid(gridTw: number[], pct: number, marginTw: number): DocxDocument {
    const cellMar =
      `<w:tblCellMar><w:left w:w="${marginTw}" w:type="dxa"/>` +
      `<w:right w:w="${marginTw}" w:type="dxa"/></w:tblCellMar>`;
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="${pct * 50}" w:type="pct"/>` +
      `<w:tblLayout w:type="fixed"/>${cellMar}</w:tblPr>` +
      `<w:tblGrid>${gridTw.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
      `<w:tr>${gridTw
        .map((w, i) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/></w:tcPr>${p(`c${i}`)}</w:tc>`)
        .join("")}</w:tr></w:tbl>`;
    return DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(tbl + SECTION),
        "word/settings.xml":
          `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          `<w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat></w:settings>`,
      }),
    );
  }

  /** Strict proportional scaling — what the columns did before the probe. */
  const proportional = (gridTw: number[], tableWidth: number): number[] => {
    const total = gridTw.reduce((a, b) => a + b, 0);
    return gridTw.map((w) => (w * tableWidth) / total);
  };

  it("scales each column's CONTENT share, not its whole width", () => {
    // parity-tables: grid 2160/813/7724 at 90% with 10pt margins. Word paints
    // 231/95/794 device px at 192 DPI. The content-share model gives
    // 231.8/95.2/796.1 over this section's 9360tw column — the residual on the
    // last column is the fixture's own column width, which back-solves to
    // ~9342tw and lands the model on 231.4/95.1/794.5. Strict proportional
    // scaling gives 226.8/85.4/811.0, out by 45px across the row, so the shape
    // is unmistakable either way.
    const grid = [2160, 813, 7724];
    const widths = renderedWidths(layout(pctGrid(grid, 90, 200)));
    expect(widths[0]).toBeCloseTo(115.92, 1);
    expect(widths[1]).toBeCloseTo(47.61, 1);
    expect(widths[2]).toBeCloseTo(398.07, 1);
    // The row still totals the table's percentage width.
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(0.9 * CONTENT_WIDTH, 2);
    // Every column is materially away from where proportional scaling put it.
    const before = proportional(grid, 0.9 * CONTENT_WIDTH);
    widths.forEach((w, i) => expect(Math.abs(w - before[i])).toBeGreaterThan(2));
  });

  it("agrees with proportional scaling when the margins are zero", () => {
    // At zero margins the two formulas are the same expression, which is why
    // the zero-margin fixtures never saw this bug.
    const grid = [2160, 813, 7724];
    const widths = renderedWidths(layout(pctGrid(grid, 90, 0)));
    const expected = proportional(grid, 0.9 * CONTENT_WIDTH);
    widths.forEach((w, i) => expect(w).toBeCloseTo(expected[i], 2));
  });

  it("splits an even grid evenly whatever the margins are", () => {
    for (const marginTw of [0, 200, 500]) {
      const widths = renderedWidths(layout(pctGrid([3000, 3000, 3000], 90, marginTw)));
      expect(widths[0]).toBeCloseTo(widths[1], 4);
      expect(widths[1]).toBeCloseTo(widths[2], 4);
      expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(0.9 * CONTENT_WIDTH, 2);
    }
  });

  it("floors a column narrower than its own margins at margins + 1pt", () => {
    // A column authored under its own margins keeps 1pt of content instead of
    // collapsing: with 400tw of margins the floor sweep plateaus at 420tw.
    const widths = renderedWidths(layout(pctGrid([100, 100, 10000], 90, 200)));
    expect(widths[0]).toBeCloseTo(420 / 15, 6);
    expect(widths[1]).toBeCloseTo(420 / 15, 6);
    // What the two floored columns take over their share comes out of the one
    // still above the floor, so the row still totals the percentage width.
    expect(widths[2]).toBeCloseTo((0.9 * 9360 - 2 * 420) / 15, 6);
    expect(widths.reduce((a, b) => a + b, 0)).toBeCloseTo(0.9 * CONTENT_WIDTH, 6);
  });

  it("raises the floor with the margins rather than by a fixed width", () => {
    // The sweep plateaus at 818tw under 800tw margins, so the floor tracks the
    // margins; margins + 20tw takes that within 2tw, a quarter of a device
    // pixel at 192 DPI.
    const widths = renderedWidths(layout(pctGrid([100, 100, 10000], 90, 400)));
    expect(widths[0]).toBeCloseTo(820 / 15, 6);
    expect(Math.abs(widths[0] - 818 / 15) * 2).toBeLessThan(1); // within 1 device px
  });
});

// ---------------------------------------------------------------------------
// Cell margins
// ---------------------------------------------------------------------------

describe("cell margins", () => {
  it("writes tblCellMar at table scope and tcMar at cell scope", () => {
    const doc = load(tableXml(2, 2));
    const caret = () => findT(docRoot(doc), "r0c0");
    expect(setTableCellMargins(doc, caret(), "table", { left: 12, right: 12 })).toBe(true);
    const tblMar = child(child(tblOf(doc), "tblPr")!, "tblCellMar")!;
    expect(tblMar.children.map((c) => localName(c.name))).toEqual(["left", "right"]);
    expect(attr(child(tblMar, "left"), "w")).toBe("240");

    expect(setTableCellMargins(doc, caret(), "cell", { top: 6, left: 30 })).toBe(true);
    const tcMar = child(tcPrOf(doc, "r0c0")!, "tcMar")!;
    expect(tcMar.children.map((c) => localName(c.name))).toEqual(["top", "left"]);
    expect(child(tcPrOf(doc, "r0c1")!, "tcMar")).toBeUndefined();
  });

  it("indents the cell's text by the override", () => {
    const doc = load(tableXml(1, 2));
    const xOf = (result: LayoutResult, text: string): number => {
      for (const page of result.pages) {
        for (const item of page.items) {
          if (item.kind === "text" && item.text.includes(text)) return item.x;
        }
      }
      throw new Error(`no text item for ${text}`);
    };
    const before = xOf(layout(doc), "r0c0");
    setTableCellMargins(doc, findT(docRoot(doc), "r0c0"), "cell", { left: 36 });
    expect(xOf(layout(doc), "r0c0")).toBeGreaterThan(before + 30);
  });

  it("drops the override so the table default applies again", () => {
    const doc = load(tableXml(1, 1));
    const caret = () => findT(docRoot(doc), "r0c0");
    setTableCellMargins(doc, caret(), "cell", { left: 36 });
    expect(setTableCellMargins(doc, caret(), "cell", null)).toBe(true);
    expect(child(tcPrOf(doc, "r0c0")!, "tcMar")).toBeUndefined();
    expect(setTableCellMargins(doc, caret(), "cell", null)).toBe(false);
  });

  it("retires a w:start twin when it sets the left inset", () => {
    const doc = load(tableXml(1, 1, `<w:tblCellMar><w:start w:w="600" w:type="dxa"/></w:tblCellMar>`));
    setTableCellMargins(doc, findT(docRoot(doc), "r0c0"), "table", { left: 5 });
    const mar = child(child(tblOf(doc), "tblPr")!, "tblCellMar")!;
    expect(mar.children.map((c) => localName(c.name))).toEqual(["left"]);
    expect(attr(child(mar, "left"), "w")).toBe("100");
  });
});

// ---------------------------------------------------------------------------
// Repeat header rows
// ---------------------------------------------------------------------------

describe("repeat header rows", () => {
  /** A table long enough to break across pages. */
  function longTable(rows: number): string {
    return tableXml(rows, 2);
  }

  const headerFlags = (doc: DocxDocument): boolean[] =>
    tblOf(doc)
      .children.filter((c) => localName(c.name) === "tr")
      .map((tr) => {
        const trPr = child(tr, "trPr");
        return !!(trPr && child(trPr, "tblHeader"));
      });

  it("marks a BAND of leading rows, not just the first", () => {
    const doc = load(longTable(6));
    expect(setTableHeaderRows(doc, tblOf(doc), 2)).toBe(true);
    expect(headerFlags(doc)).toEqual([true, true, false, false, false, false]);
    // trPr must sit first in the row and tblHeader inside it in schema order.
    const firstRow = tblOf(doc).children.filter((c) => localName(c.name) === "tr")[0];
    expect(localName(firstRow.children[0].name)).toBe("trPr");
  });

  it("shrinks the band and leaves no stranded header row", () => {
    const doc = load(longTable(6));
    setTableHeaderRows(doc, tblOf(doc), 3);
    expect(setTableHeaderRows(doc, tblOf(doc), 1)).toBe(true);
    expect(headerFlags(doc)).toEqual([true, false, false, false, false, false]);
    // Rewriting the same band changes nothing.
    expect(setTableHeaderRows(doc, tblOf(doc), 1)).toBe(false);
    expect(setTableHeaderRows(doc, tblOf(doc), 0)).toBe(true);
    expect(headerFlags(doc)).toEqual([false, false, false, false, false, false]);
  });

  it("keeps an existing trHeight and puts tblHeader after it", () => {
    const doc = load(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr><w:tc>${p("r0c0")}</w:tc></w:tr></w:tbl>`,
    );
    setTableHeaderRows(doc, tblOf(doc), 1);
    const trPr = child(tblOf(doc).children.find((c) => localName(c.name) === "tr")!, "trPr")!;
    expect(trPr.children.map((c) => localName(c.name))).toEqual(["trHeight", "tblHeader"]);
  });

  it("repeats the whole band on the continuation page", () => {
    // 90 rows overflow one page, so the table continues.
    const doc = load(longTable(90));
    setTableHeaderRows(doc, tblOf(doc), 2);
    const result = layout(doc);
    expect(result.pages.length).toBeGreaterThan(1);
    const texts = (pageIdx: number): string[] =>
      result.pages[pageIdx].items
        .filter((item) => item.kind === "text")
        .map((item) => (item.kind === "text" ? item.text : ""));
    // Both header rows are re-painted at the top of page 2.
    expect(texts(1).some((t) => t.includes("r0c0"))).toBe(true);
    expect(texts(1).some((t) => t.includes("r1c0"))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// readTableProperties — the values a Table Properties dialog prefills from
// ---------------------------------------------------------------------------

describe("readTableProperties", () => {
  const read = (doc: DocxDocument, cellText = "r0c0") =>
    readTableProperties(doc, findT(docRoot(doc), cellText));

  it("answers widths in points, from twips", () => {
    // 2 columns of 2000 twips inside a 4000-twip table: 100pt each, 200pt total.
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    const info = read(doc)!;
    expect(info.width).toEqual({ unit: "pt", value: 200 });
    expect(info.columnWidthsPt).toEqual([100, 100]);
    expect(info.columnCount).toBe(2);
  });

  it("reads a percent width back as a percentage, not as fiftieths", () => {
    const doc = load(tableXml(1, 2, `<w:tblW w:w="2500" w:type="pct"/>`));
    expect(read(doc)!.width).toEqual({ unit: "pct", value: 50 });
  });

  it("calls a table with no w:tblW auto, which is what the schema means", () => {
    const doc = load(tableXml(1, 2));
    expect(read(doc)!.width).toEqual({ unit: "auto", value: 0 });
  });

  it("reports an undeclared cell margin as ABSENT rather than as zero", () => {
    const doc = load(tableXml(1, 2));
    expect(read(doc)!.cellMargins).toEqual({});
    // A dialog can then tell "inherits the default" from "suppressed to 0".
    setTableCellMargins(doc, findT(docRoot(doc), "r0c0"), "table", { top: 6, left: 0 });
    const margins = read(doc)!.cellMargins;
    expect(margins.top).toBe(6);
    expect(margins.left).toBe(0);
    expect(margins.bottom).toBeUndefined();
  });

  it("counts the header BAND, stopping at the first row outside it", () => {
    const doc = load(tableXml(4, 2));
    expect(read(doc)!.headerRows).toBe(0);
    setTableHeaderRows(doc, tblOf(doc), 2);
    expect(read(doc)!.headerRows).toBe(2);
  });

  it("names the caret's GRID column, counting spans to its left", () => {
    expect(read(load(tableXml(1, 3)), "r0c2")!.columnIdx).toBe(2);
    // A leading cell spanning two grid columns puts the next cell at index 2,
    // which is the index setTableColumnWidth addresses.
    const spanned =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p("wide")}</w:tc>` +
      `<w:tc><w:tcPr/>${p("last")}</w:tc></w:tr></w:tbl>`;
    expect(read(load(spanned), "last")!.columnIdx).toBe(2);
  });

  it("says nothing about a caret outside a table", () => {
    const doc = load(p("plain paragraph"));
    expect(readTableProperties(doc, findT(docRoot(doc), "plain"))).toBeUndefined();
  });
});
