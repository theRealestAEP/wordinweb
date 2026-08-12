// @vitest-environment jsdom
/**
 * Mail-merge preview through the REAL component: the `mergeRecord` prop paints
 * the active record, stepping to another record repaints, and the document is
 * never touched.
 *
 * The step case is the one worth a live mount. Stepping records changes no
 * blocks, so every cache in the incremental path says "reuse is safe" — a
 * preview that keeps showing record 1 under a counter reading "2 of 2" is the
 * failure this pins, and it can only be seen end to end.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import type { DocxDocument, MergeRecord } from "@wordinweb/core";

/** A complex MERGEFIELD: begin / instrText / separate / cached / end. */
function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p>${field(`MERGEFIELD First \\b "Dear " \\f ","`, "«First»")}</w:p>` +
      `<w:p>${field("MERGEFIELD Nickname", "«Nickname»")}</w:p>` +
      `</w:body></w:document>`,
  ),
});

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  const render = async (mergeRecord?: MergeRecord) => {
    await act(async () => {
      root.render(
        createElement(DocxView, {
          source: FIXTURE,
          editable: true,
          mergeRecord,
          onReady: (api: DocxViewApi) => {
            seen.api = api;
          },
          onLoad: (info: { document: DocxDocument }) => {
            seen.doc = info.document;
          },
        }),
      );
    });
    await tick(20);
  };
  await render();
  return { container, root, seen, render };
}

describe("mail-merge preview through DocxView", () => {
  it("paints placeholders with no record and the record's values with one", async () => {
    const { container, root, render } = await mount();
    expect(container.textContent).toContain("«First»");

    await render({ First: "Alex" });
    expect(container.textContent).toContain("Dear Alex,");
    // Nickname is not in the data: its placeholder STAYS, so the user can see
    // which fields the source leaves unbound (a deliberate Word divergence).
    expect(container.textContent).toContain("«Nickname»");
    root.unmount();
  });

  it("repaints when the record steps — never the previous record's values", async () => {
    const { container, root, render } = await mount();
    await render({ First: "Alex" });
    expect(container.textContent).toContain("Dear Alex,");

    await render({ First: "Robin" });
    expect(container.textContent).toContain("Dear Robin,");
    expect(container.textContent).not.toContain("Alex");
    root.unmount();
  });

  it("suppresses \\b and \\f for a present-but-empty column", async () => {
    const { container, root, render } = await mount();
    await render({ First: "" });
    expect(container.textContent).not.toContain("Dear");
    expect(container.textContent).not.toContain(",");
    root.unmount();
  });

  it("lists the merge field names the document uses", async () => {
    const { root, seen, render } = await mount();
    await render({ First: "Alex" });
    expect(seen.api?.listMergeFieldNames()).toEqual(["First", "Nickname"]);
    root.unmount();
  });

  it("writes nothing into the document", async () => {
    const { root, seen, render } = await mount();
    const before = seen.doc!.save();
    await render({ First: "Alex" });
    expect(Buffer.from(seen.doc!.save()).equals(Buffer.from(before))).toBe(true);
    root.unmount();
  });
});
