import { describe, expect, it } from "vitest";
import { parseXml, DEFAULT_XML_LIMITS } from "../src/xml.js";

describe("parseXml resource caps (security F4)", () => {
  it("parses normal documents under the defaults", () => {
    const el = parseXml("<a><b>x</b><c/></a>");
    expect(el.name).toBe("a");
    expect(el.children.map((c) => c.name)).toEqual(["b", "c"]);
  });

  it("rejects input longer than maxLength", () => {
    expect(() => parseXml("<a/>", { ...DEFAULT_XML_LIMITS, maxLength: 3 })).toThrow(/exceeds/);
  });

  it("rejects too many nodes", () => {
    const many = "<r>" + "<x/>".repeat(50) + "</r>";
    expect(() => parseXml(many, { ...DEFAULT_XML_LIMITS, maxNodes: 10 })).toThrow(/too many nodes/);
  });

  it("rejects excessive nesting depth", () => {
    const deep = "<a>".repeat(30) + "</a>".repeat(30);
    expect(() => parseXml(deep, { ...DEFAULT_XML_LIMITS, maxDepth: 8 })).toThrow(/too deep/);
  });

  it("defaults are generous enough for a realistic document", () => {
    const doc = "<w:document><w:body>" + "<w:p><w:r><w:t>hi</w:t></w:r></w:p>".repeat(500) + "</w:body></w:document>";
    expect(() => parseXml(doc)).not.toThrow();
  });
});
