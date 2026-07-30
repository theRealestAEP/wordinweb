// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { strToU8, zipSync } from "fflate";
import { serializeXml, type DocxDocument } from "@wordinweb/core";
import { DocxView } from "../src/index.js";

function tableDoc(): Uint8Array {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  });
}

async function tick() {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 5)); });
}

describe("table text wrapping control", () => {
  it("opens from the move handle and switches between None and Around", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    let doc: DocxDocument | null = null;
    await act(async () => {
      root.render(createElement(DocxView, {
        source: tableDoc(),
        editable: true,
        onLoad: ({ document }) => { doc = document; },
      }));
    });
    for (let i = 0; i < 30 && !host.querySelector("[data-dxw-table-move]"); i++) await tick();

    const openToolbar = async () => {
      const handle = host.querySelector<HTMLElement>("[data-dxw-table-move]");
      expect(handle).toBeTruthy();
      expect(handle!.title).toContain("text wrapping options");
      await act(async () => {
        handle!.dispatchEvent(new MouseEvent("mousedown", {
          bubbles: true,
          cancelable: true,
          clientX: 20,
          clientY: 20,
          button: 0,
        }));
        document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: 20, clientY: 20 }));
      });
      expect(host.querySelector('[aria-label="Table text wrapping"]')).toBeTruthy();
    };

    await openToolbar();
    const buttons = () => [...host.querySelectorAll<HTMLButtonElement>('[aria-label="Table text wrapping"] button')];
    expect(buttons().find((button) => button.textContent === "None")?.getAttribute("aria-pressed")).toBe("true");
    act(() => { buttons().find((button) => button.textContent === "Around")!.click(); });
    await tick();
    expect(serializeXml(doc!.docRoot)).toContain("tblpPr");

    await openToolbar();
    expect(buttons().find((button) => button.textContent === "Around")?.getAttribute("aria-pressed")).toBe("true");
    act(() => { buttons().find((button) => button.textContent === "None")!.click(); });
    await tick();
    expect(serializeXml(doc!.docRoot)).not.toContain("tblpPr");

    await act(async () => { root.unmount(); });
    host.remove();
  });
});
