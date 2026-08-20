/**
 * The header/footer fast path must work on a WINDOWED document.
 *
 * Editing a header used to fall back to a full layout as soon as the page
 * window was active, because the fast path bailed on `prev._window`. On a
 * 500-page document that is the most expensive thing a header edit can do.
 *
 * The window makes this safe rather than unsafe: finalizeHeadersFooters skips
 * discarded pages, and those pages pick the new header up from the current
 * document when the window rebuilds them. These tests pin both halves — the
 * retained pages get the new header immediately, and a page rebuilt from
 * outside the window agrees with a full layout of the edited document.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, relayoutHeadersFooters } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { LayoutResult } from "../src/layout/types.js";
import type { XmlElement } from "../src/xml.js";
import { makeDocx, p, W_NS } from "./helpers.js";

const measurer = new ApproxMeasurer();

function headerPart(text: string): string {
  return `<?xml version="1.0"?><w:hdr ${W_NS}>${p(text)}</w:hdr>`;
}

function bigDocWithHeader(paras: number, header = "HEADER ORIGINAL"): DocxDocument {
  const body = Array.from({ length: paras }, (_, i) =>
    p(`para-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet`),
  ).join("");
  const documentXml =
    `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}` +
    `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>` +
    `</w:body></w:document>`;
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": documentXml,
      "word/header1.xml": headerPart(header),
      "word/_rels/document.xml.rels":
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
        `</Relationships>`,
    }),
  );
}

/** Retype the header's text in place, the way editing a header does: the
 * parsed run and its source both change, and the model generation stays put so
 * the fast path stays eligible. */
function retypeHeader(doc: DocxDocument, text: string): void {
  const header = [...doc.headers.values()][0];
  const paragraph = header.blocks[0];
  if (paragraph.type !== "paragraph") throw new Error("expected a header paragraph");
  for (const child of paragraph.children) {
    for (const run of child.type === "hyperlink" ? child.runs : [child]) {
      for (const content of run.content) {
        if (content.kind !== "text") continue;
        content.text = text;
        if (content.srcT) (content.srcT as XmlElement).text = text;
      }
    }
  }
}

/** Text of a page's header/footer band. */
function hfText(result: LayoutResult, pageIndex: number): string {
  const page = result.pages[pageIndex];
  return page.items
    .slice(page.hfStart)
    .filter((item) => item.kind === "text")
    .map((item) => (item as { text: string }).text)
    .join("");
}

describe("header/footer fast path on a windowed document", () => {
  it("retypes the header without re-laying the body", () => {
    const doc = bigDocWithHeader(2000);
    const first = layoutDocument(doc, { measurer, windowModel: true });
    expect(first._window).toBeTruthy();
    expect(hfText(first, 0)).toBe("HEADER ORIGINAL");

    retypeHeader(doc, "HEADER EDITED");
    const fast = relayoutHeadersFooters(doc, first, measurer);

    // The whole point: the window no longer forces a full layout.
    expect(fast).not.toBeNull();
    expect(fast!.pages.length).toBe(first.pages.length);
    expect(hfText(fast!, 0)).toBe("HEADER EDITED");
    // The window survives the fast path, so the document stays windowed.
    expect(fast!._window).toBeTruthy();
  });

  it("rebuilds a page from outside the window with the edited header", () => {
    const doc = bigDocWithHeader(2000);
    const first = layoutDocument(doc, { measurer, windowModel: true });
    const discarded = first.pages.findIndex((page, index) => index > 0 && page.items.length === 0);
    expect(discarded).toBeGreaterThan(0);

    retypeHeader(doc, "HEADER EDITED");
    const fast = relayoutHeadersFooters(doc, first, measurer)!;
    expect(fast).not.toBeNull();

    // A page the window had dropped rebuilds on demand, and must come back
    // with the NEW header and a body identical to a full layout's.
    fast._window!.materialize([discarded]);
    const full = layoutDocument(bigDocWithHeader(2000, "HEADER EDITED"), { measurer });
    const strip = (result: LayoutResult, index: number): string =>
      JSON.stringify(result.pages[index], (key, value) =>
        key === "src" || key === "tbl" ? undefined : value,
      );
    expect(hfText(fast, discarded)).toBe("HEADER EDITED");
    expect(strip(fast, discarded)).toBe(strip(full, discarded));
  });

  it("still refuses when the new header changes the body box", () => {
    const doc = bigDocWithHeader(2000);
    const first = layoutDocument(doc, { measurer, windowModel: true });
    // A header tall enough to push bodyTop down invalidates every retained
    // body, windowed or not.
    retypeHeader(doc, Array.from({ length: 40 }, (_, i) => `tall line ${i}`).join(" "));
    expect(relayoutHeadersFooters(doc, first, measurer)).toBeNull();
  });
});
