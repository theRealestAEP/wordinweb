import { describe, expect, it } from "vitest";
import { parseXml, sameXml, serializeXml, type XmlElement } from "../src/xml.js";

describe("sameXml", () => {
  // compareDocuments uses sameXml where it used to compare serialized
  // strings, so it must agree with serializeXml, attribute order aside.
  const x = (s: string) => parseXml(s);
  const agree = (a: XmlElement, b: XmlElement) => expect(sameXml(a, b)).toBe(serializeXml(a) === serializeXml(b));

  it("agrees with serialized equality on ordinary trees", () => {
    const p = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>`;
    agree(x(p), x(p));
    agree(x(p), x(p.replace("Hello", "Hallo")));
    agree(x(p), x(p.replace("center", "left")));
    agree(x(p), x(p.replace("<w:jc w:val=\"center\"/>", "")));
  });

  it("ignores attribute order", () => {
    expect(sameXml(x(`<w:t a="1" b="2"/>`), x(`<w:t b="2" a="1"/>`))).toBe(true);
  });

  it("skips the editor's empty-run placeholder, as serializeXml does", () => {
    const withPlaceholder = x(`<w:p><w:r><w:t>Hi</w:t></w:r><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>`);
    withPlaceholder.children[1].omitWhenEmpty = true;
    const without = x(`<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`);
    expect(serializeXml(withPlaceholder)).toBe(serializeXml(without));
    expect(sameXml(withPlaceholder, without)).toBe(true);
  });
});
