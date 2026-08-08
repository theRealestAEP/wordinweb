import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * Bibliography source management OVER THE WIRE.
 *
 * The four operations are document-scoped like the style-definition ones, and
 * the convergence story is the same: the sources part is sequenced state,
 * every value comes out of the payload, and the rejection predicate (the
 * source tag; deleteCitationSource's is-it-cited clause; setCitationStyle's
 * already-selected check) is decided from that shared state, so every replica
 * decides alike. Convergence is checked on the SAVED PACKAGE BYTES, not just
 * the XML trees, because these operations write a part save() serializes
 * separately — including the part-creation path, whose item number and
 * datastore itemID must come out identical on every replica.
 */

const SOURCES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA">` +
  `<b:Source><b:Tag>Doe03</b:Tag>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>A Study of Things</b:Title><b:Year>2003</b:Year></b:Source>` +
  `</b:Sources>`;

function makeDoc(withSources: boolean, body?: string): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body ?? `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`}` +
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

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

/** Replicated trees PLUS the saved package: the part these operations move is
 * serialized by save(), so byte-convergence is asserted there. */
function stateOf(s: DocumentSession): string {
  const xml = s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
  return xml + "||" + Buffer.from(s.doc.save()).toString("base64");
}

describe("citation source operations over the wire", () => {
  it("creates the sources part and the source identically on two replicas of a blank document", () => {
    const a = new DocumentSession(makeDoc(false));
    const b = new DocumentSession(makeDoc(false));
    const intent = {
      ...base,
      kind: "createCitationSource" as const,
      source: {
        tag: "Smith20",
        type: "book" as const,
        authors: [{ last: "Smith", first: "Ada" }],
        title: "Deep Water",
        year: "2020",
        publisher: "North Press",
      },
    };
    expect(a.submit(intent).kind).toBe("applied");
    expect(b.submit(intent).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));
    const saved = Buffer.from(a.doc.save());
    expect(saved.toString("latin1")).toContain("customXml/item1.xml");
    expect(saved.toString("latin1")).toContain("customXml/itemProps1.xml");
  });

  it("edits and deletes convergently, and rejects for an unknown tag", () => {
    const a = new DocumentSession(makeDoc(true));
    const b = new DocumentSession(makeDoc(true));
    const edit = { ...base, kind: "editCitationSource" as const, tag: "Doe03", patch: { year: "2004" } };
    expect(a.submit(edit).kind).toBe("applied");
    expect(b.submit(edit).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));

    const del = { ...base, clientSeq: 2, kind: "deleteCitationSource" as const, tag: "Doe03" };
    expect(a.submit(del).kind).toBe("applied");
    expect(b.submit(del).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));

    const before = stateOf(a);
    expect(a.submit({ ...base, clientSeq: 3, kind: "editCitationSource", tag: "Gone", patch: { year: "1" } }).kind).toBe("rejected");
    expect(a.submit({ ...base, clientSeq: 4, kind: "deleteCitationSource", tag: "Gone" }).kind).toBe("rejected");
    expect(stateOf(a)).toBe(before);
  });

  it("refuses to delete a source a CITATION field cites, on every replica alike", () => {
    const body =
      `<w:p><w:r><w:t>see</w:t></w:r>` +
      `<w:fldSimple w:instr=" CITATION Doe03 \\l 1033 "><w:r><w:t>(Doe, 2003)</w:t></w:r></w:fldSimple></w:p>`;
    const s = new DocumentSession(makeDoc(true, body));
    const before = stateOf(s);
    expect(s.submit({ ...base, kind: "deleteCitationSource", tag: "Doe03" }).kind).toBe("rejected");
    expect(stateOf(s)).toBe(before);
  });

  it("switches the citation style convergently and rejects the already-selected style", () => {
    const a = new DocumentSession(makeDoc(true));
    const b = new DocumentSession(makeDoc(true));
    const intent = { ...base, kind: "setCitationStyle" as const, style: "MLA" as const };
    expect(a.submit(intent).kind).toBe("applied");
    expect(b.submit(intent).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));
    expect(a.submit({ ...base, clientSeq: 2, kind: "setCitationStyle", style: "MLA" }).kind).toBe("rejected");
  });

  it("refuses malformed payloads before sequencing", () => {
    const cases: Record<string, unknown>[] = [
      { kind: "createCitationSource", source: { tag: "bad tag", type: "book" } },
      { kind: "createCitationSource", source: { tag: "T1", type: "novel" } },
      { kind: "createCitationSource", source: { tag: "T1", type: "book", extra: 1 } },
      { kind: "createCitationSource", source: { tag: "T1", type: "book", authors: [{ last: "" }] } },
      { kind: "editCitationSource", tag: "T1", patch: {} },
      { kind: "editCitationSource", tag: "T1", patch: { tag: "T2" } },
      { kind: "deleteCitationSource", tag: 'a"b' },
      { kind: "setCitationStyle", style: "Chicago" },
    ];
    for (const body of cases) {
      expect(validateIntent({ ...base, ...body } as never), JSON.stringify(body)).not.toBeNull();
    }
  });
});
