// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { convertEmfToDataUrl } from "emf-converter";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { paintsSettled, renderToDom } from "../src/render/dom.js";
import { makeDocxWithMedia, wrapDocument } from "./helpers.js";

vi.mock("emf-converter", () => ({ convertEmfToDataUrl: vi.fn().mockResolvedValue("data:image/png;base64,AA==") }));
afterEach(() => vi.unstubAllGlobals());

describe("EMF pictures in the browser", () => {
  it("prepares the physical frame before handing the picture to the player", async () => {
    vi.stubGlobal("URL", class extends URL { static createObjectURL = () => "blob:equation"; });
    const emf = new Uint8Array(108);
    const view = new DataView(emf.buffer);
    const header = [1, 88, 30, 40, 70, 60, 100, 200, 1100, 700, 0x464d4520, 0x10000, 108, 2, 1, 0, 0, 0, 1000, 1000, 100, 100];
    [...header, 14, 20, 0, 16, 20].forEach((value, i) => view.setInt32(i * 4, value, true));
    const doc = DocxDocument.load(makeDocxWithMedia({
      "word/document.xml": wrapDocument(`<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:shape style="width:100pt;height:50pt"><v:imagedata r:id="rId5"/></v:shape></w:pict></w:r></w:p>`),
      "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/equation.emf"/></Relationships>`,
    }, { "word/media/equation.emf": emf }));
    const container = document.createElement("div");
    renderToDom(doc, layoutDocument(doc, { measurer: new ApproxMeasurer() }), container);
    await paintsSettled();
    expect(convertEmfToDataUrl).toHaveBeenCalledOnce();
    const played = new DataView(vi.mocked(convertEmfToDataUrl).mock.calls[0][0]);
    expect([8, 12, 16, 20].map((at) => played.getInt32(at, true))).toEqual([0, 0, 100, 50]);
    expect([played.getInt32(96, true), played.getInt32(100, true)]).toEqual([-10, -20]);
    expect(container.querySelector("img")?.src).toBe("data:image/png;base64,AA==");
  });
});
