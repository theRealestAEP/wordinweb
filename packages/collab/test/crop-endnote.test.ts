import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent, DEFAULT_LIMITS } from "../src/validate.js";

/**
 * The two operations added with the crop and endnote work, over the wire.
 *
 * setCrop is the first OBJECT-addressed REGISTERED operation, so it also
 * exercises the registry's object address end to end: the carrying run's
 * stable id plus an objectIndex resolve to a drawing, and an id that names no
 * drawing rejects cleanly instead of mutating anything.
 */

const DRAWING = `<w:drawing>
  <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="914400" cy="914400"/>
    <wp:docPr id="1" name="Pic"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:blipFill><a:blip r:embed="rIdIMG" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`;

function makeDoc(body: string, extra: Record<string, Uint8Array> = {}): DocxDocument {
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
    ...extra,
  }));
}

/** A session whose first paragraph holds one text run and one picture run. */
function pictureSession(): { s: DocumentSession; runId: number; textRunId: number } {
  const doc = makeDoc(
    `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r><w:r>${DRAWING}</w:r></w:p>`,
    {
      "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdIMG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`),
      "word/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    },
  );
  const s = new DocumentSession(doc);
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const runs = para.children.filter((c) => c.type === "run") as Run[];
  return { s, runId: s.ids.idOf(runs[1].src!)!, textRunId: s.ids.idOf(runs[0].src!)! };
}

describe("setCrop intent", () => {
  it("writes a:srcRect on the drawing the addressed run carries", () => {
    const { s, runId } = pictureSession();
    const e = s.submit({
      kind: "setCrop", clientId: "a", clientSeq: 1, base: 0, runId,
      crop: { l: 0.25, t: 0.1, r: 0.05, b: 0 },
    });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain(`<a:srcRect l="25000" t="10000" r="5000"/>`);
  });

  it("resolves an explicit objectIndex to the same drawing", () => {
    const { s, runId } = pictureSession();
    const e = s.submit({
      kind: "setCrop", clientId: "a", clientSeq: 1, base: 0, runId, objectIndex: 0,
      crop: { l: 0.5, t: 0, r: 0, b: 0 },
    });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain(`<a:srcRect l="50000"/>`);
  });

  it("rejects cleanly when the addressed run carries no drawing", () => {
    const { s, textRunId } = pictureSession();
    const e = s.submit({
      kind: "setCrop", clientId: "a", clientSeq: 1, base: 0, runId: textRunId,
      crop: { l: 0.25, t: 0, r: 0, b: 0 },
    });
    expect(e.kind).toBe("rejected");
    expect(serializeXml(s.doc.docRoot)).not.toContain("srcRect");
  });

  it("applies identically on two replicas of the same document", () => {
    const build = () => {
      const { s, runId } = pictureSession();
      s.submit({ kind: "setCrop", clientId: "a", clientSeq: 1, base: 0, runId, crop: { l: 0.1, t: 0.2, r: 0.3, b: 0.05 } });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });

  it("validates the crop payload", () => {
    const base = { kind: "setCrop", clientId: "a", clientSeq: 1, base: 0, runId: 1 };
    const bad = (crop: unknown) => validateIntent({ ...base, crop } as never, DEFAULT_LIMITS);
    expect(bad({ l: 0, t: 0, r: 0, b: 0 })).toBeNull();
    expect(bad({ l: 0.4, t: 0, r: 0.4, b: 0 })).toBeNull();
    expect(bad({ l: -0.1, t: 0, r: 0, b: 0 })).toContain("bad l");
    expect(bad({ l: 1, t: 0, r: 0, b: 0 })).toContain("bad l");
    // Trimming a whole axis away leaves no picture to draw.
    expect(bad({ l: 0.6, t: 0, r: 0.5, b: 0 })).toContain("nothing left to show");
    expect(bad({ l: 0, t: 0.7, r: 0, b: 0.4 })).toContain("nothing left to show");
    expect(bad({ l: 0, t: 0, r: 0 })).toContain("bad b");
    expect(bad({ l: 0, t: 0, r: 0, b: 0, x: 1 })).toContain("unknown edge x");
    expect(bad(null)).toContain("bad crop");
    expect(
      validateIntent({ ...base, objectIndex: -1, crop: { l: 0, t: 0, r: 0, b: 0 } } as never, DEFAULT_LIMITS),
    ).toContain("bad objectIndex");
  });
});

describe("insertEndnote intent", () => {
  it("adds the reference run and the endnotes part entry", () => {
    const doc = makeDoc(`<w:p><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`);
    const s = new DocumentSession(doc);
    const run = (s.doc.sections[0].blocks[0] as Paragraph).children[0] as Run;
    const e = s.submit({
      kind: "insertEndnote", clientId: "a", clientSeq: 1, base: 0,
      runId: s.ids.idOf(run.src!)!, text: "a note", nodeIds: [700, 701, 702, 703, 704, 705, 706, 707],
    });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("endnoteReference");
    expect(s.doc.endnotes.size).toBe(1);
  });

  it("allocates the same id and writes the same XML on two replicas", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc(`<w:p><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`));
      const run = (s.doc.sections[0].blocks[0] as Paragraph).children[0] as Run;
      const runId = s.ids.idOf(run.src!)!;
      const ids = [700, 701, 702, 703, 704, 705, 706, 707];
      s.submit({ kind: "insertEndnote", clientId: "a", clientSeq: 1, base: 0, runId, text: "first", nodeIds: ids });
      s.submit({ kind: "insertEndnote", clientId: "a", clientSeq: 2, base: s.seq, runId, text: "second", nodeIds: ids.map((n) => n + 100) });
      return serializeXml(s.doc.endnotesTree()!);
    };
    const xml = build();
    expect(xml).toBe(build());
    expect(xml).toContain(`<w:endnote w:id="1">`);
    expect(xml).toContain(`<w:endnote w:id="2">`);
  });

  it("validates the note text", () => {
    const base = { kind: "insertEndnote", clientId: "a", clientSeq: 1, base: 0, runId: 1 };
    expect(validateIntent({ ...base, text: "ok" } as never, DEFAULT_LIMITS)).toBeNull();
    expect(validateIntent({ ...base, text: "   " } as never, DEFAULT_LIMITS)).toContain("empty");
    expect(validateIntent({ ...base, text: "x".repeat(20_001) } as never, DEFAULT_LIMITS)).toContain("too long");
  });
});
