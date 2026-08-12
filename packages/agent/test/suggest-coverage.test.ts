import { describe, expect, it } from "vitest";
import { DocxDocument, acceptAllRevisions, localName, rejectAllRevisions, serializeXml, type XmlElement } from "@wordinweb/core";
import { strToU8, zipSync } from "fflate";
import { AGENT_EDIT_CAPABILITIES, AgentDocument, LocalDocumentSession, type AgentBlock, type AgentOperation } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

/**
 * Every operation the capability map declares suggestable must actually write
 * a tracked form, and the reviewer must be able to undo it.
 *
 * The AI panel derives the `suggest: true` it injects from that map, so a kind
 * that advertises the flag and then edits outright is a broken promise: the
 * panel offers Accept all and Reject all over a change no revision covers.
 * A kind that advertises the flag and then REFUSES it is the same promise
 * broken from the other side — the edit fails instead of landing untracked.
 */

const TWO_PARAGRAPHS = body(
  `<w:p><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>Delta epsilon</w:t></w:r></w:p>`,
);

const WITH_BREAK = body(`<w:p><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>`);

const TABLE = body(
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>` +
  `<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
  `<w:p><w:r><w:t>After</w:t></w:r></w:p>`,
);

/** setTableStyle only takes a style the document declares, so the table seed
 * needs a styles part when that case addresses it. */
function styledTable(): Uint8Array {
  const zip = zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    ),
    "word/styles.xml": strToU8(
      `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>`,
    ),
    "word/document.xml": strToU8(TABLE),
  });
  return zip;
}

/** Any element OOXML uses to carry a revision: an insertion, a deletion, or
 * one of the five property-change records. */
const TRACKED = /<w:(?:ins|del|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange)[ />]/;

interface Refs {
  /** The first body paragraph, and the second when the seed has one. */
  first: Extract<AgentBlock, { type: "paragraph" }>;
  second?: Extract<AgentBlock, { type: "paragraph" }>;
  /** The paragraph inside the table's first cell, which addresses table ops. */
  cell: string;
}

interface Case {
  seed: string | Uint8Array;
  operation(refs: Refs): AgentOperation;
  /** Whether this document already carries the edit. Checked on the original
   * (false), the untracked edit (true), the rejected copy (false) and the
   * accepted copy (true). */
  applied(xml: string, paragraphs: string[]): boolean;
}

const CASES: Record<string, Case> = {
  insertText: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "insertText", at: { blockRef: first.ref, runRef: first.runs[0].ref, offset: 0 }, text: "New " }),
    applied: (_xml, paragraphs) => paragraphs[0] === "New Alpha beta gamma",
  },
  insertSeparator: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "insertSeparator", at: { blockRef: first.ref, runRef: first.runs[0].ref, offset: 5 }, separator: "br" }),
    applied: (xml) => xml.includes("<w:br/>"),
  },
  deleteSeparator: {
    seed: WITH_BREAK,
    operation: ({ first }) => ({ kind: "deleteSeparator", at: { blockRef: first.ref, runRef: first.runs[0].ref, offset: 5 } }),
    applied: (xml) => !xml.includes("<w:br/>"),
  },
  splitParagraph: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "splitParagraph", at: { blockRef: first.ref, runRef: first.runs[0].ref, offset: 5 } }),
    applied: (_xml, paragraphs) => paragraphs.length === 3,
  },
  mergeParagraph: {
    seed: TWO_PARAGRAPHS,
    operation: ({ second }) => ({ kind: "mergeParagraph", blockRef: second!.ref }),
    applied: (_xml, paragraphs) => paragraphs.length === 1,
  },
  formatRun: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "formatRun", blockRef: first.ref, runRef: first.runs[0].ref, patch: { bold: true } }),
    applied: (xml) => xml.includes("<w:b/>"),
  },
  formatRange: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "formatRange", blockRef: first.ref, runRef: first.runs[0].ref, start: 0, end: 5, patch: { bold: true } }),
    applied: (xml) => xml.includes("<w:b/>"),
  },
  formatParagraph: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "formatParagraph", blockRef: first.ref, styleId: "Heading1" }),
    applied: (xml) => xml.includes('w:val="Heading1"'),
  },
  setListType: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "setListType", blockRef: first.ref, listKind: "bullet" }),
    applied: (xml) => xml.includes("<w:numPr>"),
  },
  adjustIndent: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "adjustIndent", blockRef: first.ref, direction: 1 }),
    applied: (xml) => xml.includes("<w:ind "),
  },
  setSpacing: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "setSpacing", blockRef: first.ref, patch: { lineMultiple: 2 } }),
    applied: (xml) => xml.includes("<w:spacing "),
  },
  setParagraphBorders: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "setParagraphBorders", blockRef: first.ref, patch: { borders: { top: { style: "single", sz: 8 } } } }),
    applied: (xml) => xml.includes("<w:pBdr>"),
  },
  setTabStops: {
    seed: TWO_PARAGRAPHS,
    operation: ({ first }) => ({ kind: "setTabStops", blockRef: first.ref, stops: [{ posPt: 72, align: "left", leader: "none" }] }),
    applied: (xml) => xml.includes("<w:tabs>"),
  },
  "tableOp cellShading": {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "tableOp", cellRef: cell, op: { kind: "cellShading", fill: "FF0000" } }),
    applied: (xml) => xml.includes('w:fill="FF0000"'),
  },
  "tableOp cellVAlign": {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "tableOp", cellRef: cell, op: { kind: "cellVAlign", v: "center" } }),
    applied: (xml) => xml.includes("<w:vAlign "),
  },
  "tableOp textWrapping": {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "tableOp", cellRef: cell, op: { kind: "textWrapping", wrapping: "around", xPx: 10, yPx: 10 } }),
    applied: (xml) => xml.includes("<w:tblpPr "),
  },
  setTableBorders: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableBorders", cellRef: cell, scope: "table", edges: ["top"], border: { style: "single", sz: 8 } }),
    applied: (xml) => xml.includes("<w:tblBorders>"),
  },
  setTableStyle: {
    seed: styledTable(),
    operation: ({ cell }) => ({ kind: "setTableStyle", cellRef: cell, styleId: "TableGrid" }),
    applied: (xml) => xml.includes("<w:tblStyle "),
  },
  setTableLook: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableLook", cellRef: cell, look: { firstRow: true } }),
    applied: (xml) => xml.includes("<w:tblLook "),
  },
  setTableWidth: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableWidth", cellRef: cell, unit: "pct", value: 80 }),
    applied: (xml) => xml.includes('w:type="pct"'),
  },
  setTableColumnWidth: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableColumnWidth", cellRef: cell, colIdx: 0, widthPt: 120 }),
    applied: (xml) => xml.includes('<w:gridCol w:w="2400"/>'),
  },
  setTableLayout: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableLayout", cellRef: cell, layout: "fixed" }),
    applied: (xml) => xml.includes('<w:tblLayout w:type="fixed"/>'),
  },
  setTableCellMargins: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableCellMargins", cellRef: cell, scope: "table", margins: { top: 20 } }),
    applied: (xml) => xml.includes("<w:tblCellMar>"),
  },
  setTableHeaderRows: {
    seed: TABLE,
    operation: ({ cell }) => ({ kind: "setTableHeaderRows", cellRef: cell, count: 1 }),
    applied: (xml) => xml.includes("<w:tblHeader/>"),
  },
};

function load(seed: string | Uint8Array): Uint8Array {
  return typeof seed === "string" ? makeDocx(seed) : seed;
}

function connect(seed: string | Uint8Array) {
  const session = new LocalDocumentSession(load(seed));
  const agent = AgentDocument.connect(session, { provenance: { author: "AI", now: () => "2026-08-12T00:00:00Z" } });
  return { session, agent, tool: agent.tools().find((candidate) => candidate.name === "word_document_edit")! };
}

function refsOf(agent: AgentDocument): Refs {
  const read = agent.inspect({ kind: "read" });
  if (!("blocks" in read)) throw new Error("read returned no blocks");
  const paragraphs = read.blocks.filter((block) => block.type === "paragraph");
  const table = read.blocks.find((block) => block.type === "table");
  return {
    first: paragraphs[0] as Extract<AgentBlock, { type: "paragraph" }>,
    second: paragraphs[1] as Extract<AgentBlock, { type: "paragraph" }> | undefined,
    cell: table && table.type === "table" ? table.cells[0].blocks[0] : "",
  };
}

function paragraphTexts(doc: DocxDocument): string[] {
  const out: string[] = [];
  const textOf = (element: XmlElement): string =>
    (localName(element.name) === "t" ? element.text : "") + element.children.map(textOf).join("");
  const visit = (element: XmlElement): void => {
    if (localName(element.name) === "p") {
      out.push(textOf(element));
      return;
    }
    for (const child of element.children) visit(child);
  };
  visit(doc.docRoot);
  return out;
}

function carries(one: Case, doc: DocxDocument): boolean {
  return one.applied(serializeXml(doc.docRoot), paragraphTexts(doc));
}

/** Run one case on a fresh document and hand back the document it produced. */
async function apply(one: Case, suggest: boolean): Promise<DocxDocument> {
  const { session, agent, tool } = connect(one.seed);
  const operation = one.operation(refsOf(agent));
  await tool.execute({ revision: agent.revision, operations: [suggest ? { ...operation, suggest: true } : operation] });
  return session.doc;
}

function reviewed(doc: DocxDocument, resolve: (copy: DocxDocument) => void): DocxDocument {
  const copy = DocxDocument.load(doc.save());
  resolve(copy);
  return copy;
}

describe("every capability-declared suggestable kind writes a tracked form", () => {
  it("declares exactly the kinds this file covers", () => {
    const declared = Object.entries(AGENT_EDIT_CAPABILITIES)
      .filter(([, capability]) => capability.optional?.includes("suggest"))
      .map(([kind]) => kind)
      .sort();
    const covered = [...new Set(Object.keys(CASES).map((name) => name.split(" ")[0]))].sort();
    expect(covered).toEqual(declared);
  });

  for (const [name, one] of Object.entries(CASES)) {
    it(`${name} records a revision that reject undoes and accept keeps`, async () => {
      // The case has to be a real edit, or "it tracked nothing" would pass.
      expect(carries(one, DocxDocument.load(load(one.seed)))).toBe(false);
      expect(carries(one, await apply(one, false))).toBe(true);

      const doc = await apply(one, true);
      expect(serializeXml(doc.docRoot)).toMatch(TRACKED);

      const rejected = reviewed(doc, rejectAllRevisions);
      expect(serializeXml(rejected.docRoot)).not.toMatch(TRACKED);
      expect(carries(one, rejected)).toBe(false);

      const accepted = reviewed(doc, acceptAllRevisions);
      expect(serializeXml(accepted.docRoot)).not.toMatch(TRACKED);
      expect(carries(one, accepted)).toBe(true);
    });
  }
});
