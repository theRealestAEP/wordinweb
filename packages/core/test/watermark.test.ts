import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  deleteWatermark,
  headerWatermarks,
  insertWatermark,
  removeWatermark,
  setWordArtOpacity,
  setWordArtRotation,
  setWordArtText,
  wordArtOpacity,
  wordArtRotation,
  wordArtText,
} from "../src/edit/watermark.js";
import { Block, ShapeWordArt } from "../src/model.js";
import { XmlElement, localName } from "../src/xml.js";
import { W_NS, makeDocx } from "./helpers.js";

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

// ---------------------------------------------------------------------------
// Authoring
// ---------------------------------------------------------------------------

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W_NS}><w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p></w:hdr>`;

/** A package with `count` header parts, all referenced from the one section. */
function docWithHeaders(count: number): Uint8Array {
  const parts: Record<string, string> = {};
  const types = ["default", "first", "even"];
  const refs: string[] = [];
  const rels: string[] = [];
  const overrides: string[] = [];
  for (let i = 0; i < count; i++) {
    parts[`word/header${i + 1}.xml`] = HEADER_XML;
    refs.push(`<w:headerReference w:type="${types[i]}" r:id="rIdH${i + 1}"/>`);
    rels.push(
      `<Relationship Id="rIdH${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header${i + 1}.xml"/>`,
    );
    overrides.push(
      `<Override PartName="/word/header${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
    );
  }
  parts["word/_rels/document.xml.rels"] =
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
  parts["[Content_Types].xml"] = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${overrides.join("")}
</Types>`;
  parts["word/document.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
  <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
  <w:sectPr>${refs.join("")}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
</w:body></w:document>`;
  return makeDocx(parts);
}

/** Every wordart shape the PARSER found in the header models — the check that
 * matters, because it is the same path the renderer draws from. */
function parsedHeaderWordArt(doc: DocxDocument): ShapeWordArt[] {
  const found: ShapeWordArt[] = [];
  const visit = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") continue;
      for (const child of block.children) {
        if (child.type !== "run") continue;
        for (const content of child.content) {
          if (content.kind === "anchor" && content.shape.type === "wordart") found.push(content.shape);
        }
      }
    }
  };
  for (const hf of doc.headers.values()) visit(hf.blocks);
  return found;
}

describe("watermark authoring", () => {
  it("stamps a washed-out diagonal watermark into every header part", () => {
    const doc = DocxDocument.load(docWithHeaders(3));
    expect(insertWatermark(doc, { text: "DRAFT" })).toBe(true);

    const shapes = parsedHeaderWordArt(doc);
    expect(shapes).toHaveLength(3);
    for (const shape of shapes) {
      expect(shape.text).toBe("DRAFT");
      expect(shape.rotation).toBe(315);
      expect(shape.opacity).toBe(0.5);
      expect(shape.behind).toBe(true);
      expect(shape.hAlign).toBe("center");
      expect(shape.vAlign).toBe("center");
      expect(shape.hRel).toBe("margin");
      expect(shape.vRel).toBe("margin");
      // The whole point of emitting the full _x0000_t136 guide path: a
      // degenerate one makes Word paint the literal 1pt text instead of
      // fitting the glyphs to the box.
      expect(shape.noFit).toBeUndefined();
      expect(shape.width).toBeGreaterThan(shape.height);
    }
  });

  it("survives save and reload as the same watermark", () => {
    const doc = DocxDocument.load(docWithHeaders(1));
    insertWatermark(doc, { text: "CONFIDENTIAL", color: "FF0000", opacity: 0.25 });
    const reloaded = DocxDocument.load(doc.save());

    const shapes = parsedHeaderWordArt(reloaded);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].text).toBe("CONFIDENTIAL");
    expect(shapes[0].fill).toBe("#FF0000");
    expect(shapes[0].opacity).toBe(0.25);
    expect(shapes[0].noFit).toBeUndefined();
  });

  it("lays the text flat when diagonal is false", () => {
    const doc = DocxDocument.load(docWithHeaders(1));
    insertWatermark(doc, { text: "SAMPLE", diagonal: false });
    const shape = parsedHeaderWordArt(doc)[0];
    expect(shape.rotation).toBe(0);
    // Flat text has to sit inside the 6.5in text column of a Letter page.
    expect(shape.width).toBeLessThanOrEqual(6.5 * 96);
  });

  it("replaces the watermark rather than stacking a second one", () => {
    const doc = DocxDocument.load(docWithHeaders(2));
    insertWatermark(doc, { text: "DRAFT" });
    insertWatermark(doc, { text: "FINAL" });
    const shapes = parsedHeaderWordArt(doc);
    expect(shapes).toHaveLength(2);
    expect(shapes.map((s) => s.text)).toEqual(["FINAL", "FINAL"]);
  });

  it("gives the shapes in different header parts distinct, non-random ids", () => {
    const first = DocxDocument.load(docWithHeaders(3));
    const second = DocxDocument.load(docWithHeaders(3));
    insertWatermark(first, { text: "DRAFT" });
    insertWatermark(second, { text: "DRAFT" });

    const ids = headerWatermarks(first).map((shape) => shape.attrs["id"]);
    expect(new Set(ids).size).toBe(3);
    // Two replicas of the same room must author byte-identical XML.
    expect(headerWatermarks(second).map((shape) => shape.attrs["id"])).toEqual(ids);
  });

  it("removes the watermark and leaves the header's own content alone", () => {
    const doc = DocxDocument.load(docWithHeaders(2));
    insertWatermark(doc, { text: "DRAFT" });
    expect(removeWatermark(doc)).toBe(true);
    expect(parsedHeaderWordArt(doc)).toHaveLength(0);
    // Nothing left to remove is a clean no-op, not a failure to report.
    expect(removeWatermark(doc)).toBe(false);

    const reloaded = DocxDocument.load(doc.save());
    expect(parsedHeaderWordArt(reloaded)).toHaveLength(0);
    for (const hf of reloaded.headers.values()) expect(hf.blocks).toHaveLength(1);
  });

  it("declines a document with no header part", () => {
    const doc = DocxDocument.load(docWithHeaders(0));
    expect(insertWatermark(doc, { text: "DRAFT" })).toBe(false);
    expect(removeWatermark(doc)).toBe(false);
  });

  it("leaves parts it did not touch byte-identical", () => {
    const original = docWithHeaders(2);
    const doc = DocxDocument.load(original);
    insertWatermark(doc, { text: "DRAFT" });

    const before = unzipSync(original);
    const after = unzipSync(doc.save());
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    for (const name of Object.keys(before)) {
      if (name.startsWith("word/header")) continue;
      expect(after[name], name).toEqual(before[name]);
    }
  });
});
