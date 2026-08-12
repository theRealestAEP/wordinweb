import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  DocxDocument,
  insertShapeAt,
  operationBody,
  serializeXml,
  type Paragraph,
  type Run,
  type TextContent,
} from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * setDrawingTextFit OVER THE WIRE.
 *
 * OBJECT-addressed like setCrop and setModel3DRotation: the carrying run's
 * stable id plus which of its contents the drawing is. The whole payload is
 * the mode and an optional cached percentage — nothing is read from a local
 * layout — so two replicas that apply the same intent write the same bytes.
 */

function makeDoc(): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">Anchor</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
    }),
  );
}

/** A document whose first paragraph already carries a text box, as an opened
 * file would. insertShapeAt authors <a:noAutofit/>. */
function docWithShape(): DocxDocument {
  const doc = makeDoc();
  const para = doc.sections[0].blocks[0] as Paragraph;
  const t = ((para.children[0] as Run).content[0] as TextContent).srcT!;
  expect(insertShapeAt(doc, t, "textBox", "Overfull shape text")).not.toBeNull();
  return doc;
}

/** The stable id of the run holding the shape, and its index within that run. */
function address(session: DocumentSession): { runId: number; objectIndex: number } {
  for (const block of session.doc.sections[0].blocks) {
    if (block.type !== "paragraph") continue;
    for (const child of block.children) {
      if (child.type !== "run") continue;
      const objectIndex = child.content.findIndex((content) => content.kind === "anchor");
      if (objectIndex >= 0) return { runId: session.ids.idOf(child.src!)!, objectIndex };
    }
  }
  throw new Error("no shape in the document");
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

function fitIntent(session: DocumentSession, args: { mode: string; fontScalePct?: number }) {
  const { runId, objectIndex } = address(session);
  return { ...base, ...operationBody("setDrawingTextFit", runId, { objectIndex, ...args } as never) };
}

function xmlOf(session: DocumentSession): string {
  return session.doc.editableRoots().map((root) => serializeXml(root)).join("|");
}

describe("setDrawingTextFit over the wire", () => {
  it("lands the same bodyPr on two replicas, byte for byte", () => {
    const a = new DocumentSession(docWithShape());
    const b = new DocumentSession(docWithShape());
    const intent = fitIntent(a, { mode: "resizeShape" });
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain("<a:spAutoFit/>");
    expect(xmlOf(a)).not.toContain("<a:noAutofit/>");
  });

  it("writes the cached scale in ECMA-376 thousandths of a percent", () => {
    const a = new DocumentSession(docWithShape());
    const b = new DocumentSession(docWithShape());
    const intent = fitIntent(a, { mode: "shrinkText", fontScalePct: 62.5 });
    a.submit(intent as never);
    b.submit(intent as never);
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain('<a:normAutofit fontScale="62500"/>');
  });

  it("omits the scale when none is given, which is what Word writes", () => {
    // probe-shapefit: Word leaves a bare <a:normAutofit/> bare — it computes
    // no fontScale of its own on the file path.
    const session = new DocumentSession(docWithShape());
    session.submit(fitIntent(session, { mode: "shrinkText" }) as never);
    expect(xmlOf(session)).toContain("<a:normAutofit/>");
  });

  it("replaces the autofit already there rather than stacking a second one", () => {
    const session = new DocumentSession(docWithShape());
    session.submit(fitIntent(session, { mode: "resizeShape" }) as never);
    session.submit({ ...fitIntent(session, { mode: "none" }), clientSeq: 2 } as never);
    const xml = xmlOf(session);
    expect(xml).toContain("<a:noAutofit/>");
    expect(xml).not.toContain("<a:spAutoFit/>");
    expect(xml.match(/Autofit/g)).toHaveLength(1);
  });

  it("rejects an unresolvable object address instead of guessing", () => {
    const session = new DocumentSession(docWithShape());
    const { runId } = address(session);
    const unknownRun = { ...base, ...operationBody("setDrawingTextFit", 999999, { objectIndex: 0, mode: "resizeShape" } as never) };
    expect(session.submit(unknownRun as never).kind).toBe("rejected");
    const pastTheEnd = { ...base, ...operationBody("setDrawingTextFit", runId, { objectIndex: 99, mode: "resizeShape" } as never) };
    expect(session.submit(pastTheEnd as never).kind).toBe("rejected");
  });

  it("is a clean no-op on a drawing with no bodyPr to write into", () => {
    const session = new DocumentSession(makeDoc());
    const para = session.doc.sections[0].blocks[0] as Paragraph;
    const run = para.children[0] as Run;
    const intent = {
      ...base,
      ...operationBody("setDrawingTextFit", session.ids.idOf(run.src!)!, { objectIndex: 0, mode: "resizeShape" } as never),
    };
    // Content 0 of that run is plain text, so the address resolves to no
    // drawing at all.
    expect(session.submit(intent as never).kind).toBe("rejected");
  });

  it("refuses a malformed payload before it reaches the document", () => {
    const cases: [string, unknown][] = [
      ["a missing mode", { objectIndex: 0 }],
      ["an unknown mode", { objectIndex: 0, mode: "shrinkBox" }],
      ["a scale below the schema range", { objectIndex: 0, mode: "shrinkText", fontScalePct: 0 }],
      ["a scale above the schema range", { objectIndex: 0, mode: "shrinkText", fontScalePct: 101 }],
      ["a non-numeric scale", { objectIndex: 0, mode: "shrinkText", fontScalePct: "62.5" }],
      // The scale is a:normAutofit's cache and belongs to no other mode.
      ["a scale on resizeShape", { objectIndex: 0, mode: "resizeShape", fontScalePct: 62.5 }],
    ];
    for (const [why, payload] of cases) {
      const intent = { ...base, kind: "setDrawingTextFit", runId: 1, ...(payload as object) };
      expect(validateIntent(intent as never), why).not.toBeNull();
    }
    expect(validateIntent({ ...base, kind: "setDrawingTextFit", runId: 1, mode: "none" } as never)).toBeNull();
  });
});
