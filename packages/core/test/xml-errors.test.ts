import { describe, expect, it } from "vitest";
import { asPdfRenderError } from "../src/pdf/errors.js";
import { DEFAULT_XML_LIMITS, XmlParseError, parseXml, rejectXmlDeclarations } from "../src/xml.js";

describe("XML parse errors", () => {
  it.each([
    "<!DOCTYPE x><x/>",
    '<!DOCTYPE x [<!ENTITY a "expanded">]><x>&a;</x>',
    '<!ENTITY a "expanded"><x/>',
    "<?xml",
    "<!--",
    "<![CDATA[",
    "</x>",
    "<x",
    '<x a="',
    "",
    "<x>&#x110000;</x>",
  ])("classifies %s as damaged", (xml) => {
    let error: unknown;
    try {
      parseXml(xml);
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(XmlParseError);
    expect(asPdfRenderError(error)).toMatchObject({ code: "damaged", cause: error });
  });

  it("classifies each parser resource limit as damaged", () => {
    for (const limits of [{ maxLength: 1 }, { maxNodes: 1 }, { maxDepth: 1 }]) {
      expect(() => parseXml("<x><y/></x>", { ...DEFAULT_XML_LIMITS, ...limits })).toThrow(XmlParseError);
    }
  });

  it("preserves declaration-like text in comments, CDATA and processing instructions", () => {
    const xml = '<?example <!DOCTYPE x> ?><!-- <!ENTITY a "text"> --><x><![CDATA[<!DOCTYPE x>]]>&amp;&#65;</x>';
    expect(() => rejectXmlDeclarations(xml)).not.toThrow();
    expect(parseXml(xml).text).toBe("<!DOCTYPE x>&A");
  });
});
