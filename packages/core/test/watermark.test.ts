import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  deleteWatermark,
  setWordArtOpacity,
  setWordArtRotation,
  setWordArtText,
  wordArtOpacity,
  wordArtRotation,
  wordArtText,
} from "../src/edit/watermark.js";
import { XmlElement, localName } from "../src/xml.js";
import { makeDocx } from "./helpers.js";

const VML_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"';

function docWithWatermark(text = "CONFIDENTIAL", style = "position:absolute;margin-left:0;margin-top:0;width:400pt;height:100pt;rotation:315;z-index:-1"): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${VML_NS}><w:body>
  <w:p><w:r><w:pict>
    <v:shape id="wm1" type="#_x0000_t136" style="${style}" fillcolor="#c0c0c0">
      <v:fill opacity="0.5"/>
      <v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="${text}"/>
    </v:shape>
  </w:pict></w:r></w:p>
  <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
</w:body></w:document>`;
}

function loadDoc(documentXml: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": documentXml }));
}

/** Find the v:shape element in the retained document tree. */
function findShape(doc: DocxDocument): XmlElement | undefined {
  const root = (doc as unknown as { docRoot: XmlElement }).docRoot;
  let found: XmlElement | undefined;
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "shape") found = el;
    for (const c of el.children) walk(c);
  };
  walk(root);
  return found;
}

/** Serialize + reparse, returning the reloaded shape (round-trip through save). */
function roundTrip(doc: DocxDocument): { doc: DocxDocument; shape: XmlElement | undefined } {
  const bytes = doc.save();
  const reloaded = DocxDocument.load(bytes);
  return { doc: reloaded, shape: findShape(reloaded) };
}

describe("watermark editing", () => {
  it("reads the current text, opacity, and rotation", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    expect(shape).toBeTruthy();
    expect(wordArtText(shape)).toBe("CONFIDENTIAL");
    expect(wordArtOpacity(shape)).toBeCloseTo(0.5, 5);
    expect(wordArtRotation(shape)).toBe(315);
  });

  it("edits the watermark text and it persists through save/reload", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    expect(setWordArtText(doc, shape, "DRAFT")).toBe(true);
    expect(wordArtText(shape)).toBe("DRAFT");
    const rt = roundTrip(doc);
    expect(wordArtText(rt.shape!)).toBe("DRAFT");
  });

  it("escapes special characters in the watermark text", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    setWordArtText(doc, shape, 'A & B < "C"');
    const rt = roundTrip(doc);
    expect(wordArtText(rt.shape!)).toBe('A & B < "C"');
  });

  it("sets opacity, clamps to 0..1, and persists", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    expect(setWordArtOpacity(doc, shape, 0.25)).toBe(true);
    expect(wordArtOpacity(shape)).toBeCloseTo(0.25, 5);
    setWordArtOpacity(doc, shape, 5);
    expect(wordArtOpacity(shape)).toBe(1);
    const rt = roundTrip(doc);
    expect(wordArtOpacity(rt.shape!)).toBe(1);
  });

  it("creates a v:fill when the shape has none, to hold opacity", () => {
    // A shape without a <v:fill> child (fill comes only from fillcolor attr).
    const xml = `<?xml version="1.0"?>
<w:document ${VML_NS}><w:body>
  <w:p><w:r><w:pict>
    <v:shape id="wm2" type="#_x0000_t136" style="position:absolute;width:400pt;height:100pt;z-index:-1" fillcolor="#808080">
      <v:textpath string="TOP SECRET"/>
    </v:shape>
  </w:pict></w:r></w:p>
</w:body></w:document>`;
    const doc = loadDoc(xml);
    const shape = findShape(doc)!;
    expect(wordArtOpacity(shape)).toBe(1); // no v:fill -> default opaque
    setWordArtOpacity(doc, shape, 0.4);
    expect(wordArtOpacity(shape)).toBeCloseTo(0.4, 5);
    const rt = roundTrip(doc);
    expect(wordArtOpacity(rt.shape!)).toBeCloseTo(0.4, 5);
  });

  it("sets rotation, normalizes to [0,360), and persists", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    setWordArtRotation(doc, shape, 45);
    expect(wordArtRotation(shape)).toBe(45);
    setWordArtRotation(doc, shape, 405); // -> 45
    expect(wordArtRotation(shape)).toBe(45);
    setWordArtRotation(doc, shape, -45); // -> 315
    expect(wordArtRotation(shape)).toBe(315);
    const rt = roundTrip(doc);
    expect(wordArtRotation(rt.shape!)).toBe(315);
  });

  it("removes the rotation declaration when set to 0", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    setWordArtRotation(doc, shape, 0);
    expect(shape.attrs["style"]).not.toContain("rotation");
    expect(wordArtRotation(shape)).toBe(0);
  });

  it("deletes the watermark and its enclosing run", () => {
    const doc = loadDoc(docWithWatermark());
    const shape = findShape(doc)!;
    expect(deleteWatermark(doc, shape)).toBe(true);
    expect(findShape(doc)).toBeUndefined();
    const rt = roundTrip(doc);
    expect(rt.shape).toBeUndefined();
    // The body text paragraph must survive.
    const root = (rt.doc as unknown as { docRoot: XmlElement }).docRoot;
    let hasBody = false;
    const walk = (el: XmlElement): void => {
      if (localName(el.name) === "t" && el.text === "Body text") hasBody = true;
      for (const c of el.children) walk(c);
    };
    walk(root);
    expect(hasBody).toBe(true);
  });
});
