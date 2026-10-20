// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, layoutDocumentAsync } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, W_NS } from "./helpers.js";

// The editor's view of section-page-headers.test.ts's first case: an empty
// titlePg section 1, then a continuous section 2 with its own headers. Page 1
// is section 1's first page, so the DOM must paint S1's first-page header.

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PAGE = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;

function load(): DocxDocument {
  const parts: Record<string, string> = {};
  const rels: string[] = [];
  for (const id of ["S1FIRST", "S2DEFAULT"]) {
    parts[`word/${id}.xml`] = `<w:hdr ${W_NS}><w:p><w:r><w:t>${id}HEADER</w:t></w:r></w:p></w:hdr>`;
    rels.push(`<Relationship Id="${id}" Type="${REL}/header" Target="${id}.xml"/>`);
  }
  let body = `<w:p><w:pPr><w:sectPr><w:headerReference w:type="first" r:id="S1FIRST"/>${PAGE}<w:titlePg/></w:sectPr></w:pPr></w:p>`;
  for (let k = 1; k <= 120; k++) body += `<w:p><w:r><w:t>Body ${k}.</w:t></w:r></w:p>`;
  body += `<w:sectPr><w:headerReference w:type="default" r:id="S2DEFAULT"/><w:type w:val="continuous"/>${PAGE}</w:sectPr>`;
  parts["word/document.xml"] = `<w:document ${W_NS} xmlns:r="${REL}"><w:body>${body}</w:body></w:document>`;
  parts["word/_rels/document.xml.rels"] =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
  return DocxDocument.load(makeDocx(parts));
}

function pageTexts(container: HTMLElement): string[] {
  return [...container.querySelectorAll<HTMLElement>(".dxw-page")].map((page) => page.textContent ?? "");
}

describe("DOM renderer: the page an empty section opened", () => {
  it("paints section 1's first-page header on page 1, in the editor's layouts", async () => {
    const doc = load();
    const measurer = new ApproxMeasurer();
    const layouts = [
      layoutDocument(doc, { measurer, windowModel: true }),
      await layoutDocumentAsync(doc, { measurer }),
    ];
    for (const layout of layouts) {
      const container = document.createElement("div");
      renderToDom(doc, layout, container);
      const pages = pageTexts(container);
      expect(pages.length).toBeGreaterThan(1);
      expect(pages[0]).toContain("S1FIRSTHEADER");
      expect(pages[0]).toContain("Body 1.");
      expect(pages[0]).not.toContain("S2DEFAULTHEADER");
      for (const page of pages.slice(1)) expect(page).toContain("S2DEFAULTHEADER");
    }
  });
});
