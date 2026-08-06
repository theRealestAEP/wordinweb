import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  DocxDocument,
  insertModel3DAt,
  operationBody,
  serializeXml,
  type Paragraph,
  type Run,
  type TextContent,
} from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * setModel3DRotation OVER THE WIRE — and the registry's first OBJECT address.
 *
 * A 3D model already present in an opened file stays draggable in a room: the
 * `model3D` toolbar flag hides INSERTION only. The drag used to write the new
 * orientation into the local document and emit nothing, so it forked the room
 * silently — the same shape of bug insertImage once had.
 *
 * The operation is addressed by OBJECT: the stable id of the carrier run plus
 * which of that run's contents the drawing is. Both halves have to resolve, so
 * it keeps the honest-no-op predicate every addressed operation has, and the
 * angles ride in the payload because only the originator saw the drag.
 */

const GLB = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);
const POSTER = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

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

/** A document whose first paragraph already carries a 3D model, as an opened
 * file would. Insertion is deliberately NOT done over the wire here — this
 * operation is about editing a model that is already in the document. */
function docWithModel(): DocxDocument {
  const doc = makeDoc();
  const para = doc.sections[0].blocks[0] as Paragraph;
  const t = ((para.children[0] as Run).content[0] as TextContent).srcT!;
  expect(insertModel3DAt(doc, t, { data: GLB, poster: POSTER })).toBe(true);
  return doc;
}

/** The stable id of the run holding the model, and its index within that run. */
function address(s: DocumentSession): { runId: number; objectIndex: number } {
  for (const block of s.doc.sections[0].blocks) {
    if (block.type !== "paragraph") continue;
    for (const child of block.children) {
      if (child.type !== "run") continue;
      const objectIndex = child.content.findIndex((c) => c.kind === "image" && c.model3D);
      if (objectIndex >= 0) return { runId: s.ids.idOf(child.src!)!, objectIndex };
    }
  }
  throw new Error("no 3D model in the document");
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

function rotateIntent(s: DocumentSession, rotation = { x: 30, y: 45, z: 0 }) {
  const { runId, objectIndex } = address(s);
  return { ...base, ...operationBody("setModel3DRotation", runId, { objectIndex, rotation } as never) };
}

function xmlOf(s: DocumentSession): string {
  return s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
}

describe("setModel3DRotation over the wire", () => {
  it("lands the same orientation on two replicas, byte for byte", () => {
    const a = new DocumentSession(docWithModel());
    const b = new DocumentSession(docWithModel());
    const intent = rotateIntent(a);
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    // Degrees become DrawingML 60000ths: 30deg = 1800000, 45deg = 2700000.
    expect(xmlOf(a)).toContain('ax="1800000"');
    expect(xmlOf(a)).toContain('ay="2700000"');
  });

  it("normalizes a negative angle the same way on every replica", () => {
    const a = new DocumentSession(docWithModel());
    const b = new DocumentSession(docWithModel());
    const intent = rotateIntent(a, { x: -90, y: 720, z: 0 });
    a.submit(intent as never);
    b.submit(intent as never);
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain('ax="16200000"'); // -90 -> 270deg
    expect(xmlOf(a)).toContain('ay="0"'); // 720 -> 0deg
  });

  it("rejects an unresolvable object address instead of guessing", () => {
    // Both halves of the address must resolve. A run id nobody has, and a real
    // run whose contents have moved, are each a clean rejection everywhere
    // rather than a mutation of whatever now sits at that index.
    const s = new DocumentSession(docWithModel());
    const { runId } = address(s);
    const rotation = { x: 10, y: 10, z: 10 };
    const unknownRun = { ...base, ...operationBody("setModel3DRotation", 999999, { objectIndex: 0, rotation } as never) };
    expect(s.submit(unknownRun as never).kind).toBe("rejected");
    const pastTheEnd = { ...base, ...operationBody("setModel3DRotation", runId, { objectIndex: 99, rotation } as never) };
    expect(s.submit(pastTheEnd as never).kind).toBe("rejected");
    expect(xmlOf(s)).not.toContain('ax="600000"');
  });

  it("is a clean no-op on a drawing that is not a 3D model", () => {
    const s = new DocumentSession(makeDoc());
    const para = s.doc.sections[0].blocks[0] as Paragraph;
    const run = para.children[0] as Run;
    const intent = {
      ...base,
      ...operationBody("setModel3DRotation", s.ids.idOf(run.src!)!, {
        objectIndex: 0,
        rotation: { x: 1, y: 2, z: 3 },
      } as never),
    };
    // Content 0 of that run is plain text, so the address resolves to no
    // drawing at all.
    expect(s.submit(intent as never).kind).toBe("rejected");
  });

  it("refuses a malformed payload before it reaches the document", () => {
    const cases: [string, unknown][] = [
      ["a missing rotation", undefined],
      ["a non-object rotation", 45],
      ["a missing axis", { x: 1, y: 2 }],
      ["a non-numeric axis", { x: 1, y: 2, z: "3" }],
      ["an infinite axis", { x: 1, y: 2, z: Number.POSITIVE_INFINITY }],
    ];
    for (const [why, rotation] of cases) {
      const intent = { ...base, kind: "setModel3DRotation", runId: 1, objectIndex: 0, rotation };
      expect(validateIntent(intent as never), why).not.toBeNull();
    }
  });
});
