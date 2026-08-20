/**
 * #154: the table move handle has to arm when the pointer is ON it, not only
 * after the pointer has crossed the table.
 *
 * The handle sits 24px OUTSIDE the table's top-left corner and starts life
 * with `opacity: 0; pointer-events: none`, lifted by a mousemove handler on
 * the page surface. That handler armed only on "pointer inside the table", so
 * a pointer arriving from the toolbar or the margin — straight to where the
 * handle is — found a square that was invisible, unclickable, and therefore
 * showed neither its move cursor nor its "Move table" title. It advertised
 * itself to people who had already hovered the table, and to nobody else.
 *
 * This lives in the react package because it needs a DOM and core's suite
 * runs in node. It drives the same surface mousemove a real pointer does.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { afterEach, describe, expect, it } from "vitest";
import { DocxView, type DocxViewApi } from "../src/index.js";

const CELL = (t: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p><w:r><w:t xml:space="preserve">Before</w:t></w:r></w:p>` +
      `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
      `<w:tr>${CELL("r0c0")}${CELL("r0c1")}</w:tr><w:tr>${CELL("r1c0")}${CELL("r1c1")}</w:tr></w:tbl>` +
      `<w:p><w:r><w:t xml:space="preserve">After</w:t></w:r></w:p></w:body></w:document>`,
  ),
});

let root: Root | null = null;
let host: HTMLElement | null = null;

afterEach(async () => {
  if (root) await act(async () => root!.unmount());
  host?.remove();
  root = null;
  host = null;
});

async function mountWithTable(): Promise<HTMLElement> {
  host = document.createElement("div");
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () => {
    root!.render(createElement(DocxView, { source: FIXTURE, editable: true, onReady: (_: DocxViewApi) => {} }));
  });
  for (let i = 0; i < 40 && !host.querySelector("[data-dxw-table-move]"); i++) {
    await act(async () => { await new Promise((r) => setTimeout(r, 5)); });
  }
  const handle = host.querySelector<HTMLElement>("[data-dxw-table-move]");
  if (!handle) throw new Error("no table move handle rendered");
  return handle;
}

/** A mousemove at a point in the surface's own layout coordinates. */
function pointAt(surface: HTMLElement, x: number, y: number) {
  const rect = surface.getBoundingClientRect();
  surface.dispatchEvent(new MouseEvent("mousemove", {
    bubbles: true, clientX: rect.left + x, clientY: rect.top + y,
  }));
}

describe("#154 · the table move handle can be reached", () => {
  it("arms when the pointer lands on the handle without crossing the table", async () => {
    const handle = await mountWithTable();
    const surface = handle.parentElement!;
    const left = parseFloat(handle.style.left);
    const top = parseFloat(handle.style.top);
    expect(Number.isFinite(left) && Number.isFinite(top), "the handle has no position").toBe(true);

    // Straight onto the handle's own 22x22 square. Nothing has hovered the
    // table, which is the whole point.
    pointAt(surface, left + 11, top + 11);

    expect(handle.style.opacity, "the handle stayed invisible under the pointer").toBe("1");
    expect(handle.style.pointerEvents, "the handle still refuses the press it advertises").toBe("auto");
  });

  it("still arms the old way, by crossing the table", async () => {
    const handle = await mountWithTable();
    const surface = handle.parentElement!;
    const left = parseFloat(handle.style.left);
    const top = parseFloat(handle.style.top);
    // The table's top-left is 24px in from the handle's.
    pointAt(surface, left + 24 + 40, top + 24 + 20);
    expect(handle.style.opacity, "hovering the table no longer reveals the handle").toBe("1");
  });
});
