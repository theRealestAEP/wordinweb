import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, operationBody, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * insertBibliography / refreshBibliography OVER THE WIRE.
 *
 * A bibliography's entries are a pure function of the sources part —
 * sequenced state — so unlike a TOC's page numbers they are DERIVED on each
 * replica rather than carried, and unlike a TOC rebuild the refresh rides
 * the wire. `entryCount` is the insertToc id-budget pattern; convergence is
 * asserted on both the XML and the stable-id assignment, because a diverging
 * id table breaks every later addressed edit.
 */

const SOURCES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA">` +
  `<b:Source><b:Tag>Doe03</b:Tag>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Jane</b:First></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>A Study of Things</b:Title><b:Year>2003</b:Year><b:Publisher>Acme</b:Publisher></b:Source>` +
  `</b:Sources>`;

function makeDoc(): DocxDocument {
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
      "word/styles.xml": strToU8(
        `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`,
      ),
      "customXml/item1.xml": strToU8(SOURCES_XML),
    }),
  );
}

function anchorRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

function xmlOf(s: DocumentSession): string {
  return s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
}

function runIds(s: DocumentSession): (number | undefined)[] {
  const out: (number | undefined)[] = [];
  const walk = (el: { name: string; children: unknown[] }): void => {
    if (el.name === "w:r" || el.name.endsWith(":r")) out.push(s.ids.idOf(el as never));
    for (const c of el.children) walk(c as never);
  };
  for (const root of s.doc.editableRoots()) walk(root as never);
  return out;
}

function insertIntent(s: DocumentSession) {
  let next = 900;
  return {
    ...base,
    ...operationBody(
      "insertBibliography",
      anchorRunId(s),
      { entryCount: 1 },
      (n) => Array.from({ length: n }, () => next++),
    ),
  };
}

describe("insertBibliography over the wire", () => {
  it("derives the same entries and carried ids on two replicas from the shared sources part", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = insertIntent(a);
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(runIds(a)).toEqual(runIds(b));
    expect(xmlOf(a)).toContain(" BIBLIOGRAPHY ");
    expect(xmlOf(a)).toContain("Doe, J. (2003). A Study of Things. Acme.");
  });

  it("refuses a malformed entryCount before sequencing", () => {
    const s = new DocumentSession(makeDoc());
    for (const entryCount of [0, -1, 1.5, 10001]) {
      const intent = { ...insertIntent(s), entryCount };
      expect(validateIntent(intent as never), String(entryCount)).not.toBeNull();
    }
  });
});

describe("refreshBibliography over the wire", () => {
  it("rebuilds convergently after a replicated source edit, and rejects when already current", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const insert = insertIntent(a);
    expect(a.submit(insert as never).kind).toBe("applied");
    expect(b.submit(insert as never).kind).toBe("applied");

    // Already current: the honest no-op, decided identically on a replica
    // that never mutated anything.
    let next = 950;
    const alloc = (n: number) => Array.from({ length: n }, () => next++);
    const staleRefresh = { ...base, clientSeq: 2, kind: "refreshBibliography" as const, entryCount: 1, nodeIds: alloc(10) };
    expect(a.submit(staleRefresh).kind).toBe("rejected");
    expect(b.submit(staleRefresh).kind).toBe("rejected");

    const edit = { ...base, clientSeq: 3, kind: "editCitationSource" as const, tag: "Doe03", patch: { year: "2005" } };
    expect(a.submit(edit).kind).toBe("applied");
    expect(b.submit(edit).kind).toBe("applied");

    next = 970;
    const refresh = { ...base, clientSeq: 4, kind: "refreshBibliography" as const, entryCount: 1, nodeIds: alloc(10) };
    expect(a.submit(refresh).kind).toBe("applied");
    expect(b.submit(refresh).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(runIds(a)).toEqual(runIds(b));
    expect(xmlOf(a)).toContain("Doe, J. (2005). A Study of Things. Acme.");
    expect(xmlOf(a)).not.toContain("(2003). A Study of Things");
  });
});
