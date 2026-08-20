import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, operationBody, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * insertMergeField / insertCitation OVER THE WIRE.
 *
 * Both are run-addressed registered operations, so the stable id is the whole
 * honest-no-op story for insertMergeField. insertCitation adds one predicate
 * of its own: the tag must name a source in the package's sources part. Every
 * replica holds the same part at the intent's sequenced position, so the
 * lookup — and the display text derived from it — agrees everywhere without
 * carrying either on the wire.
 */

const SOURCES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA">` +
  `<b:Source><b:Tag>Doe03</b:Tag>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>A Study of Things</b:Title><b:Year>2003</b:Year></b:Source>` +
  `</b:Sources>`;

function makeDoc(withSources = true): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>` +
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
      ...(withSources
        ? {
            "word/_rels/document.xml.rels": strToU8(
              `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`,
            ),
            "customXml/item1.xml": strToU8(SOURCES_XML),
          }
        : {}),
    }),
  );
}

/** The stable id of the run in the anchor paragraph. */
function anchorRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

function intentFor(s: DocumentSession, kind: "insertMergeField" | "insertCitation", args: Record<string, unknown>) {
  let next = 900;
  return {
    ...base,
    ...operationBody(kind, anchorRunId(s), args as never, (n) => Array.from({ length: n }, () => next++)),
  };
}

function xmlOf(s: DocumentSession): string {
  return s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
}

/** Every run's stable id, in document order. */
function runIds(s: DocumentSession): (number | undefined)[] {
  const out: (number | undefined)[] = [];
  const walk = (el: { name: string; children: unknown[] }): void => {
    if (el.name === "w:r" || el.name.endsWith(":r")) out.push(s.ids.idOf(el as never));
    for (const c of el.children) walk(c as never);
  };
  for (const root of s.doc.editableRoots()) walk(root as never);
  return out;
}

describe("insertMergeField over the wire", () => {
  it("inserts the same placeholder field on two replicas, byte for byte", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = intentFor(a, "insertMergeField", { name: "First Name" });
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(runIds(a)).toEqual(runIds(b));
    expect(xmlOf(a)).toContain(`w:instr=" MERGEFIELD &quot;First Name&quot; \\* MERGEFORMAT "`);
    expect(xmlOf(a)).toContain("«First Name»");
  });

  it("refuses a malformed name before it reaches the document", () => {
    const s = new DocumentSession(makeDoc());
    for (const name of ["", 'a"b', "a\\b", "x".repeat(65), "café"]) {
      const intent = intentFor(s, "insertMergeField", { name });
      expect(validateIntent(intent as never), JSON.stringify(name)).not.toBeNull();
    }
  });
});

describe("insertCitation over the wire", () => {
  it("derives the same display text on both replicas from the shared sources part", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = intentFor(a, "insertCitation", { tag: "Doe03" });
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(runIds(a)).toEqual(runIds(b));
    expect(xmlOf(a)).toContain(`w:instr=" CITATION Doe03 \\l 1033 "`);
    expect(xmlOf(a)).toContain("(Doe, 2003)");
  });

  it("is a clean no-op for a tag the sources part does not hold", () => {
    const s = new DocumentSession(makeDoc());
    const before = xmlOf(s);
    expect(s.submit(intentFor(s, "insertCitation", { tag: "Nope99" }) as never).kind).toBe("rejected");
    expect(xmlOf(s)).toBe(before);
  });

  it("is a clean no-op when the document has no sources part at all", () => {
    const s = new DocumentSession(makeDoc(false));
    const before = xmlOf(s);
    expect(s.submit(intentFor(s, "insertCitation", { tag: "Doe03" }) as never).kind).toBe("rejected");
    expect(xmlOf(s)).toBe(before);
  });

  it("refuses a malformed tag before it reaches the document", () => {
    const s = new DocumentSession(makeDoc());
    for (const tag of ["", "a b", 'a"b', "a\\b", "x".repeat(65)]) {
      const intent = intentFor(s, "insertCitation", { tag });
      expect(validateIntent(intent as never), JSON.stringify(tag)).not.toBeNull();
    }
  });
});
