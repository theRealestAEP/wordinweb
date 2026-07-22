import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { StableIds } from "../src/edit/ids.js";
import { addComment, replyToComment } from "../src/edit/comments.js";
import { recordedProvenance } from "../src/edit/provenance.js";
import { serializeXml, XmlElement, localName } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import type { Paragraph, Run, TextContent } from "../src/model.js";
import type { SelectionSegment } from "../src/edit/commands.js";

function loadDoc(body: string, extra: Record<string, string> = {}) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...extra }));
}

const THREE_PARAS = p("alpha") + p("beta") + p("gamma");

function segFor(doc: DocxDocument, blockIdx: number, start: number, end: number): SelectionSegment {
  const para = doc.sections[0].blocks[blockIdx] as Paragraph;
  const run = para.children[0] as Run;
  const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT;
  return { run, t: t as SelectionSegment["t"], start, end, props: run.props };
}

describe("StableIds", () => {
  it("assigns ids in document order, deterministically across loads", () => {
    const a = loadDoc(THREE_PARAS);
    const b = loadDoc(THREE_PARAS);
    const idsA = new StableIds();
    const idsB = new StableIds();
    idsA.assignFromRoots(a.editableRoots());
    idsB.assignFromRoots(b.editableRoots());
    expect(idsA.size()).toBe(idsB.size());
    expect(idsA.size()).toBe(6); // 3 paragraphs + 3 runs
    // Same walk order → same id for the structurally same element.
    const parasA = a.sections[0].blocks as Paragraph[];
    const parasB = b.sections[0].blocks as Paragraph[];
    for (let i = 0; i < 3; i++) {
      expect(idsA.idOf(parasA[i].src)).toBe(idsB.idOf(parasB[i].src));
    }
  });

  it("keeps ids for surviving elements across refresh and re-assign", () => {
    const doc = loadDoc(THREE_PARAS);
    const ids = new StableIds();
    ids.assignFromRoots(doc.editableRoots());
    const beta = (doc.sections[0].blocks[1] as Paragraph).src;
    const betaId = ids.idOf(beta)!;
    // Mutate text in place (identity-preserving edit) and refresh.
    const run = (doc.sections[0].blocks[1] as Paragraph).children[0] as Run;
    const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
    t.text = "beta-edited";
    doc.refresh();
    ids.assignFromRoots(doc.editableRoots());
    expect(ids.idOf(beta)).toBe(betaId);
    expect((doc.sections[0].blocks[1] as Paragraph).src).toBe(beta);
  });

  it("installs carried ids for new nodes and rejects collisions", () => {
    const doc = loadDoc(THREE_PARAS);
    const ids = new StableIds();
    ids.assignFromRoots(doc.editableRoots());
    const fresh: XmlElement = { name: "w:p", attrs: {}, children: [], text: "" };
    const carried = 5000;
    expect(ids.assign(fresh, carried)).toBe(carried);
    expect(ids.elOf(carried)).toBe(fresh);
    expect(ids.nextId()).toBe(5001);
    const dup: XmlElement = { name: "w:p", attrs: {}, children: [], text: "" };
    expect(() => ids.assign(dup, carried)).toThrow(/already in use/);
  });

  it("prunes deleted elements and never reuses their ids", () => {
    const doc = loadDoc(THREE_PARAS);
    const ids = new StableIds();
    ids.assignFromRoots(doc.editableRoots());
    const before = ids.size();
    const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const beta = (doc.sections[0].blocks[1] as Paragraph).src!;
    const betaId = ids.idOf(beta)!;
    body.children.splice(body.children.indexOf(beta), 1);
    doc.refresh();
    ids.prune(doc.editableRoots());
    expect(ids.size()).toBe(before - 2); // paragraph + its run
    expect(ids.elOf(betaId)).toBeUndefined();
    const fresh: XmlElement = { name: "w:p", attrs: {}, children: [], text: "" };
    expect(ids.assign(fresh)).toBeGreaterThan(betaId);
  });

  it("encodes a caret as {blockId, runId, offset}", () => {
    const doc = loadDoc(THREE_PARAS);
    const ids = new StableIds();
    ids.assignFromRoots(doc.editableRoots());
    const para = doc.sections[0].blocks[0] as Paragraph;
    const run = para.children[0] as Run;
    const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
    const enc = ids.encodeCaret(t, 3, (el) => doc.findParentOf(el));
    expect(enc).not.toBeNull();
    expect(enc!.blockId).toBe(ids.idOf(para.src));
    expect(enc!.runId).toBe(ids.idOf(run.src));
    expect(enc!.offset).toBe(3);
    // Decode side: elOf resolves back to the same elements.
    expect(ids.elOf(enc!.blockId)).toBe(para.src);
    expect(ids.elOf(enc!.runId)).toBe(run.src);
  });
});

describe("round-trip byte identity with ids assigned (Phase 1 gate)", () => {
  it("open → assign ids → save leaves every part byte-identical", () => {
    const bytes = makeDocx({ "word/document.xml": wrapDocument(THREE_PARAS) });
    const doc = DocxDocument.load(bytes);
    const ids = new StableIds();
    ids.assignFromRoots(doc.editableRoots());
    const saved = doc.save();
    const before = unzipSync(bytes);
    const after = unzipSync(saved);
    // IDs live only in the side table: no part content may change.
    for (const [name, content] of Object.entries(before)) {
      expect(strFromU8(after[name] ?? new Uint8Array())).toBe(strFromU8(content));
    }
  });
});

describe("deterministic provenance (Phase 1 determinism fixes)", () => {
  it("addComment with recorded provenance produces identical XML on two replicas", () => {
    const run = () => {
      const doc = loadDoc(THREE_PARAS);
      const prov = recordedProvenance({
        dates: ["2026-07-22T12:00:00.000Z"],
        paraIds: ["0A1B2C3D"],
      });
      const ok = addComment(doc, [segFor(doc, 0, 0, 5)], "note", "Alex", "A", prov);
      expect(ok).toBe(true);
      return doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    };
    expect(run()).toBe(run());
  });

  it("replyToComment threads with recorded provenance deterministically", () => {
    const run = () => {
      const doc = loadDoc(THREE_PARAS);
      const prov = recordedProvenance({
        dates: ["2026-07-22T12:00:00.000Z", "2026-07-22T12:01:00.000Z"],
        paraIds: ["0A1B2C3D", "0E1F2A3B"],
      });
      addComment(doc, [segFor(doc, 0, 0, 5)], "note", "Alex", "A", prov);
      const parent = doc.comments[0];
      const ok = replyToComment(doc, parent.id, "reply", "Sam", "S", prov);
      expect(ok).toBe(true);
      return doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    };
    expect(run()).toBe(run());
  });

  it("nextDrawingId seeds past existing docPr ids instead of colliding", () => {
    const drawing =
      `<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:inline><wp:extent cx="914400" cy="914400"/>` +
      `<wp:docPr id="1500" name="Picture 1"/>` +
      `</wp:inline></w:drawing></w:r></w:p>`;
    const doc = loadDoc(drawing + p("text"));
    expect(doc.nextDrawingId()).toBe(1501);
    expect(doc.nextDrawingId()).toBe(1502);
    // Fresh document with no drawings keeps the historical seed.
    const fresh = loadDoc(p("only text"));
    expect(fresh.nextDrawingId()).toBe(1000);
  });
});

describe("save() is side-effect-free (N1 checkpoint-purity invariant)", () => {
  const pctTable =
    `<w:tbl><w:tblPr><w:tblW w:w="100%" w:type="pct"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="5"/><w:gridCol w:w="5"/></w:tblGrid>` +
    `<w:tr>` +
    `<w:tc><w:p><w:r><w:t xml:space="preserve">left cell content</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:p><w:r><w:t xml:space="preserve">right</w:t></w:r></w:p></w:tc>` +
    `</w:tr></w:tbl>`;

  it("leaves the live tree byte-identical across repeated saves (percentage grid)", () => {
    const doc = loadDoc(pctTable + p("after"));
    const before = doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    const s1 = doc.save();
    const afterFirst = doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    expect(afterFirst).toBe(before); // save did not mutate authoritative state
    const s2 = doc.save();
    // Deterministic checkpoint output: same live state → same bytes.
    expect(strFromU8(unzipSync(s2)["word/document.xml"])).toBe(
      strFromU8(unzipSync(s1)["word/document.xml"]),
    );
  });

  it("leaves content-types byte-identical when a custom-properties Override would be added", () => {
    const customProps =
      `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ` +
      `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(p("x")), "docProps/custom.xml": customProps }),
    );
    const beforeCt = serializeXml((doc as unknown as { contentTypesRoot: XmlElement }).contentTypesRoot ?? { name: "", attrs: {}, children: [], text: "" });
    const saved = doc.save();
    const afterCt = serializeXml((doc as unknown as { contentTypesRoot: XmlElement }).contentTypesRoot ?? { name: "", attrs: {}, children: [], text: "" });
    expect(afterCt).toBe(beforeCt); // live content-types not mutated
    // But the SAVED bytes do contain the Override.
    expect(strFromU8(unzipSync(saved)["[Content_Types].xml"])).toContain("/docProps/custom.xml");
  });
});

describe("DocxDocument stable-id lifecycle (opt-in)", () => {
  it("is null by default (zero cost for local-only docs)", () => {
    const doc = loadDoc(THREE_PARAS);
    expect(doc.stableIds).toBeNull();
    doc.refresh();
    expect(doc.stableIds).toBeNull(); // refresh does nothing when disabled
  });

  it("maintains the table across refresh once enabled: survivors keep ids, new nodes get ids, deleted retire", () => {
    const doc = loadDoc(THREE_PARAS);
    const ids = doc.enableStableIds();
    const beta = (doc.sections[0].blocks[1] as Paragraph).src;
    const betaId = ids.idOf(beta)!;
    const sizeBefore = ids.size();

    // In-place text edit + refresh: beta keeps its id.
    const run = (doc.sections[0].blocks[1] as Paragraph).children[0] as Run;
    const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
    t.text = "beta!";
    doc.refresh();
    expect(ids.idOf(beta)).toBe(betaId);
    expect(ids.size()).toBe(sizeBefore);

    // Delete gamma + refresh: its id retires, survivors unchanged.
    const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const gamma = (doc.sections[0].blocks[2] as Paragraph).src!;
    const gammaId = ids.idOf(gamma)!;
    body.children.splice(body.children.indexOf(gamma), 1);
    doc.refresh();
    expect(ids.elOf(gammaId)).toBeUndefined();
    expect(ids.idOf(beta)).toBe(betaId);
  });
});
