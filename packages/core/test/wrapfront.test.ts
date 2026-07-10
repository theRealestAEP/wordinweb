import { describe, expect, it } from "vitest";
import { setImageWrap } from "../src/edit/images.js";
import { child, localName, attr } from "../src/xml.js";

function fakeDoc() {
  return { refresh() {} } as any;
}
const el = (name: string, attrs: Record<string, string> = {}, children: any[] = []) => ({ name, attrs, children, text: "" });

describe("setImageWrap none/behind", () => {
  it("behind then none round-trips behindDoc", () => {
    const extent = el("wp:extent", { cx: "914400", cy: "914400" });
    const graphic = el("a:graphic");
    const anchor = el("wp:anchor", { behindDoc: "0" }, [extent, el("wp:wrapSquare", { wrapText: "bothSides" }), el("wp:docPr", { id: "1", name: "img" }), graphic]);
    const drawing = el("w:drawing", {}, [anchor]);
    expect(setImageWrap(fakeDoc(), drawing as any, "behind")).toBe(true);
    expect(attr(anchor as any, "behindDoc")).toBe("1");
    expect(anchor.children.some((c: any) => localName(c.name) === "wrapNone")).toBe(true);
    expect(setImageWrap(fakeDoc(), drawing as any, "none")).toBe(true);
    expect(attr(anchor as any, "behindDoc")).toBe("0");
    expect(anchor.children.some((c: any) => localName(c.name) === "wrapNone")).toBe(true);
    expect(anchor.children.filter((c: any) => localName(c.name).startsWith("wrap")).length).toBe(1);
  });
});
