import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { parseXml, child, attr } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";

describe("xml parser", () => {
  it("parses elements, attributes, entities", () => {
    const root = parseXml(`<?xml version="1.0"?><w:p w:val="a&amp;b"><w:t>hi &lt;there&gt;</w:t></w:p>`);
    expect(root.name).toBe("w:p");
    expect(attr(root, "val")).toBe("a&b");
    expect(child(root, "t")?.text).toBe("hi <there>");
  });

  it("handles self-closing tags and CDATA", () => {
    const root = parseXml(`<a><b/><c><![CDATA[x<y]]></c></a>`);
    expect(root.children.length).toBe(2);
    expect(child(root, "c")?.text).toBe("x<y");
  });
});

describe("document parsing", () => {
  it("decodes private-use text written in the legacy Symbol font", () => {
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>` +
          `<w:t>\uF067\uF020\uF071\uF02E</w:t></w:r></w:p>`,
      ),
    }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error("expected paragraph");
    const run = para.children[0];
    if (run.type !== "run") throw new Error("expected run");
    expect(run.content[0]).toMatchObject({ kind: "text", text: "γ θ." });
  });

  it("parses the VML preview inside an embedded OLE object", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:t>before</w:t></w:r><w:r><w:object>
            <v:shape style="width:56.5pt;height:21.4pt">
              <v:imagedata r:id="rId7"/>
            </v:shape>
          </w:object></w:r><w:r><w:t>after</w:t></w:r></w:p>`,
        ),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/equation.wmf"/>
</Relationships>`,
      }),
    );
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error("expected paragraph");
    const objectRun = para.children[1];
    if (objectRun.type !== "run") throw new Error("expected run");
    // VML pict extents round to whole points (56.5 -> 57pt, 21.4 -> 21pt):
    // Word's PDF draws every wild2-math-eq-as-images equation raster at
    // integer pt (31.45->31, 49.65->50, 120.75->121, 290.75->291).
    expect(objectRun.content[0]).toMatchObject({
      kind: "image",
      part: "word/media/equation.wmf",
      width: 76,
      height: 28,
    });
  });

  it("uses a QUOTE field's raster preview instead of duplicating its OLE object", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:r><w:object><v:shape style="width:200pt;height:40pt"><v:imagedata r:id="rId7"/></v:shape></w:object></w:r>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> QUOTE </w:instrText></w:r>
            <w:r><w:pict><v:shape style="width:170pt;height:25pt"><v:imagedata r:id="rId8"/></v:shape></w:pict></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:pict><v:shape style="width:190pt;height:30pt"><v:imagedata r:id="rId9"/></v:shape></w:pict></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>`,
        ),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/equation.wmf"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stale.png"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/equation.png"/>
</Relationships>`,
      }),
    );
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error("expected paragraph");
    const images = para.children.flatMap((child) =>
      child.type === "run" ? child.content.filter((content) => content.kind === "image") : [],
    );
    expect(images).toHaveLength(1);
    expect(images[0]).toMatchObject({ part: "word/media/equation.png" });
  });

  it("preserves the OLE object when a QUOTE field has no picture result", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:r><w:object><v:shape style="width:200pt;height:40pt"><v:imagedata r:id="rIdOutside"/></v:shape></w:object></w:r>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> QUOTE </w:instrText></w:r>
            <w:r><w:pict><v:shape style="width:170pt;height:25pt"><v:imagedata r:id="rIdStale1"/></v:shape></w:pict></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>
          <w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> QUOTE </w:instrText></w:r>
            <w:r><w:pict><v:shape style="width:170pt;height:25pt"><v:imagedata r:id="rIdStale2"/></v:shape></w:pict></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:object><v:shape style="width:190pt;height:30pt"><v:imagedata r:id="rIdResult"/></v:shape></w:object></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>`,
        ),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOutside" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/outside.wmf"/>
  <Relationship Id="rIdStale1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stale1.png"/>
  <Relationship Id="rIdStale2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stale2.png"/>
  <Relationship Id="rIdResult" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/result.wmf"/>
</Relationships>`,
      }),
    );
    const outsidePara = doc.sections[0].blocks[0];
    const resultPara = doc.sections[0].blocks[1];
    if (outsidePara.type !== "paragraph" || resultPara.type !== "paragraph") {
      throw new Error("expected paragraphs");
    }
    const outsideImages = outsidePara.children.flatMap((child) =>
      child.type === "run" ? child.content.filter((content) => content.kind === "image") : [],
    );
    const resultImages = resultPara.children.flatMap((child) =>
      child.type === "run" ? child.content.filter((content) => content.kind === "image") : [],
    );
    expect(outsideImages).toHaveLength(1);
    expect(outsideImages[0]).toMatchObject({ part: "word/media/outside.wmf" });
    expect(resultImages).toHaveLength(1);
    expect(resultImages[0]).toMatchObject({ part: "word/media/result.wmf" });
  });

  it("parses paragraphs and runs with formatting", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Bold text</w:t></w:r></w:p>`,
        ),
      }),
    );
    expect(doc.sections.length).toBe(1);
    const para = doc.sections[0].blocks[0];
    expect(para.type).toBe("paragraph");
    if (para.type !== "paragraph") return;
    const run = para.children[0];
    expect(run.type).toBe("run");
    if (run.type !== "run") return;
    expect(run.props.bold).toBe(true);
    expect(run.props.size).toBeCloseTo((14 * 4) / 3, 3); // 28 half-points = 14pt
    expect(run.content[0]).toMatchObject({ kind: "text", text: "Bold text" });
  });

  it("parses section properties (page size, margins)", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          p("hello") +
            `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="850" w:header="708" w:footer="708"/></w:sectPr>`,
        ),
      }),
    );
    const sp = doc.sections[0].props;
    expect(sp.pageWidth).toBeCloseTo((11906 / 20) * (4 / 3), 1); // A4 width
    expect(sp.pageHeight).toBeCloseTo((16838 / 20) * (4 / 3), 1);
    expect(sp.marginTop).toBeCloseTo((1134 / 20) * (4 / 3), 1);
  });

  it("parses complex fields (PAGE) across runs", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>7</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>`,
        ),
      }),
    );
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error("expected paragraph");
    const fields = para.children.flatMap((c) =>
      c.type === "run" ? c.content.filter((x) => x.kind === "field") : [],
    );
    expect(fields.length).toBe(1);
    expect(fields[0]).toMatchObject({ instruction: " PAGE ", cachedResult: "7" });
  });

  it("parses tables with grid and spans", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:tbl>
            <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>` + p("after"),
        ),
      }),
    );
    const tbl = doc.sections[0].blocks[0];
    expect(tbl.type).toBe("table");
    if (tbl.type !== "table") return;
    expect(tbl.grid.length).toBe(2);
    expect(tbl.rows[0].cells[0].props.gridSpan).toBe(2);
    expect(tbl.rows[1].cells.length).toBe(2);
  });

  it("parses bottom-to-top table cell text direction", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:tcPr><w:textDirection w:val="btLr"/></w:tcPr>${p("up")}</w:tc>
              <w:tc><w:tcPr><w:textDirection w:val="tbRl"/></w:tcPr>${p("other")}</w:tc>
            </w:tr>
          </w:tbl>`,
        ),
      }),
    );
    const table = doc.sections[0].blocks[0];
    if (table.type !== "table") throw new Error("expected table");

    expect(table.rows[0].cells[0].props.textDirection).toBe("btLr");
    expect(table.rows[0].cells[1].props.textDirection).toBeUndefined();
  });

  it("resolves styles through basedOn chains", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>`,
        ),
        "word/styles.xml": `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Base">
    <w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Child">
    <w:basedOn w:val="Base"/>
    <w:rPr><w:color w:val="0000FF"/></w:rPr>
  </w:style>
</w:styles>`,
      }),
    );
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error("expected paragraph");
    const run = para.children[0];
    if (run.type !== "run") throw new Error("expected run");
    const props = doc.effectiveRunProps(para, run.props);
    expect(props.bold).toBe(true); // inherited from Base
    expect(props.color).toBe("#0000FF"); // overridden by Child
  });

  it("resolves bidi theme fonts without changing the Latin or East Asian channels", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:rtl/></w:rPr><w:t>text</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:cstheme="minorBidi"/><w:rtl/></w:rPr><w:t>more</w:t></w:r></w:p>`,
        ),
        "word/settings.xml": `<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:themeFontLang w:bidi="he-IL"/>
</w:settings>`,
        "word/theme/theme1.xml": `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements><a:fontScheme name="Test">
    <a:majorFont>
      <a:latin typeface="Cambria"/><a:ea typeface="MS Mincho"/><a:cs typeface=""/>
      <a:font script="Hebr" typeface="Times New Roman"/>
    </a:majorFont>
    <a:minorFont>
      <a:latin typeface="Calibri"/><a:ea typeface="MS Gothic"/><a:cs typeface="Tahoma"/>
      <a:font script="Hebr" typeface="Arial"/>
    </a:minorFont>
  </a:fontScheme><a:clrScheme name="Test"/></a:themeElements>
</a:theme>`,
      }),
    );
    const first = doc.sections[0].blocks[0];
    const second = doc.sections[0].blocks[1];
    if (first.type !== "paragraph" || second.type !== "paragraph") throw new Error("expected paragraphs");
    const firstRun = first.children[0];
    const secondRun = second.children[0];
    if (firstRun.type !== "run" || secondRun.type !== "run") throw new Error("expected runs");

    const firstProps = doc.effectiveRunProps(first, firstRun.props);
    const secondProps = doc.effectiveRunProps(second, secondRun.props);
    expect(firstProps.font).toBe("Cambria");
    expect(firstProps.fontEastAsia).toBe("Cambria");
    expect(firstProps.fontComplex).toBe("Times New Roman");
    expect(secondProps.fontComplex).toBe("Tahoma");

    const saved = DocxDocument.load(doc.save()).pkg.text("word/document.xml");
    expect(saved).toContain('w:cstheme="majorBidi"');
    expect(saved).not.toContain('w:cs="Times New Roman"');
  });

  it("suppresses generated TOC hyperlink styling and sizes cached tabs from the paragraph mark", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
            <w:hyperlink w:anchor="_Toc1">
              <w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:sz w:val="36"/></w:rPr><w:t>Generated entry</w:t></w:r>
              <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:tab/></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>4</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:hyperlink>
          </w:p>
          <w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>
            <w:hyperlink w:anchor="_Toc2">
              <w:r><w:t>Marked entry</w:t></w:r>
              <w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:tab/></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc2 \\h </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>5</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:hyperlink>
          </w:p>
          <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
            <w:hyperlink w:anchor="manual">
              <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Manual link</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
            <w:hyperlink w:anchor="_Toc3">
              <w:r><w:rPr><w:rStyle w:val="Hyperlink-toc"/></w:rPr><w:t>Custom blue entry</w:t></w:r>
              <w:r><w:tab/></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> PAGEREF _Toc3 \\h </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>6</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:hyperlink>
          </w:p>`,
        ),
        "word/styles.xml": `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
  <w:style w:type="paragraph" w:styleId="TOC1"><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:rPr><w:rFonts w:ascii="Arial"/><w:i/><w:color w:val="0000FF"/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink-toc">
    <w:name w:val="Hyperlink-toc"/><w:rPr><w:color w:val="0000FF"/></w:rPr>
  </w:style>
</w:styles>`,
      }),
    );
    const generated = doc.sections[0].blocks[0];
    const marked = doc.sections[0].blocks[1];
    const manual = doc.sections[0].blocks[2];
    const custom = doc.sections[0].blocks[3];
    if (
      generated.type !== "paragraph" ||
      marked.type !== "paragraph" ||
      manual.type !== "paragraph" ||
      custom.type !== "paragraph"
    ) {
      throw new Error("expected paragraphs");
    }
    const generatedLink = generated.children[0];
    const markedLink = marked.children[0];
    const manualLink = manual.children[0];
    const customLink = custom.children[0];
    if (
      generatedLink.type !== "hyperlink" ||
      markedLink.type !== "hyperlink" ||
      manualLink.type !== "hyperlink" ||
      customLink.type !== "hyperlink"
    ) {
      throw new Error("expected hyperlinks");
    }

    const generatedProps = doc.effectiveRunProps(generated, generatedLink.runs[0].props);
    const tabProps = doc.effectiveRunProps(generated, generatedLink.runs[1].props);
    const pageProps = doc.effectiveRunProps(generated, generatedLink.runs.at(-1)!.props);
    expect(generatedProps).toMatchObject({ font: "Arial", color: "auto", underline: "none" });
    expect(generatedProps.size).toBe(24);
    expect(generatedProps.italic).toBeUndefined();
    expect(tabProps.font).toBe("Calibri");
    expect(tabProps.size).toBeCloseTo((10 * 4) / 3);
    expect(pageProps.font).toBe("Calibri");

    const markedTabProps = doc.effectiveRunProps(marked, markedLink.runs[1].props);
    expect(markedTabProps.size).toBeCloseTo((11 * 4) / 3);

    const manualProps = doc.effectiveRunProps(manual, manualLink.runs[0].props);
    expect(manualProps).toMatchObject({ font: "Arial", size: 16, italic: true });

    const customProps = doc.effectiveRunProps(custom, customLink.runs[0].props);
    const customPageProps = doc.effectiveRunProps(custom, customLink.runs.at(-1)!.props);
    expect(customProps).toMatchObject({ color: "#0000FF", underline: "none" });
    expect(customPageProps).toMatchObject({ color: "auto" });
  });

  it("parses headers and footers", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          p("body") +
            `<w:sectPr>
              <w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rId5"/>
              <w:pgSz w:w="12240" w:h="15840"/>
            </w:sectPr>`,
        ),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
        "word/header1.xml": `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${p("My Header")}</w:hdr>`,
      }),
    );
    expect(doc.headers.size).toBe(1);
    expect(doc.sections[0].props.headerRefs.default).toBe("rId5");
    const hdr = doc.headers.get("rId5")!;
    expect(hdr.blocks.length).toBe(1);
  });

  it("parses review comments and maps their anchor ranges", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:r><w:t>before </w:t></w:r>
            <w:commentRangeStart w:id="7"/>
            <w:r><w:t>flagged</w:t></w:r>
            <w:r><w:t> words</w:t></w:r>
            <w:commentRangeEnd w:id="7"/>
            <w:r><w:commentReference w:id="7"/></w:r>
            <w:r><w:t> after</w:t></w:r>
          </w:p>` +
            `<w:p>
            <w:r><w:t>point anchor</w:t></w:r>
            <w:r><w:commentReference w:id="8"/></w:r>
          </w:p>`,
        ),
        "word/comments.xml": `<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="7" w:author="Ada" w:initials="A" w:date="2026-06-01T10:00:00Z">
    <w:p><w:r><w:t>First line.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second line.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="8" w:author="Bob">
    <w:p><w:r><w:t>Point comment.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`,
      }),
    );
    expect(doc.comments.length).toBe(2);
    expect(doc.comments[0]).toMatchObject({ id: "7", author: "Ada", initials: "A" });
    expect(doc.comments[0].text).toBe("First line.\nSecond line.");
    expect(doc.comments[1]).toMatchObject({ id: "8", author: "Bob", text: "Point comment." });

    const anchors = doc.commentAnchors();
    const ranged = anchors.get("7")!;
    expect(ranged.map((t) => t.text)).toEqual(["flagged", " words"]);
    // Point comment (no range) anchors to the nearest preceding w:t.
    expect(anchors.get("8")!.map((t) => t.text)).toEqual(["point anchor"]);
  });

  it("deletes a comment with undo and save round-trip", async () => {
    const { deleteComment } = await import("../src/edit/comments.js");
    const { EditHistory } = await import("../src/edit/history.js");
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:commentRangeStart w:id="7"/>
            <w:r><w:t>flagged</w:t></w:r>
            <w:commentRangeEnd w:id="7"/>
            <w:r><w:commentReference w:id="7"/></w:r>
          </w:p>`,
        ),
        "word/comments.xml": `<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="7" w:author="Ada"><w:p><w:r><w:t>Note.</w:t></w:r></w:p></w:comment>
</w:comments>`,
      }),
    );
    const history = new EditHistory(doc);
    expect(doc.comments.length).toBe(1);

    history.checkpoint();
    expect(deleteComment(doc, "7")).toBe(true);
    expect(doc.comments.length).toBe(0);
    expect(doc.commentAnchors().size).toBe(0);
    // Markers are gone from the document part; text survives.
    const saved = DocxDocument.load(doc.save());
    expect(saved.comments.length).toBe(0);
    expect(saved.pkg.text("word/document.xml")).not.toContain("commentRangeStart");
    expect(saved.pkg.text("word/document.xml")).toContain("flagged");

    // Undo restores both the markers and the comment body.
    expect(history.undo()).toBe(true);
    expect(doc.comments.length).toBe(1);
    expect(doc.commentAnchors().get("7")!.map((t) => t.text)).toEqual(["flagged"]);
  });

  it("replies thread under the parent and round-trip; thread delete cascades", async () => {
    const { replyToComment, deleteComment } = await import("../src/edit/comments.js");
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p>
            <w:commentRangeStart w:id="0"/>
            <w:r><w:t>flagged</w:t></w:r>
            <w:commentRangeEnd w:id="0"/>
            <w:r><w:commentReference w:id="0"/></w:r>
          </w:p>`,
        ),
        "word/comments.xml": `<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Ada"><w:p><w:r><w:t>Parent note.</w:t></w:r></w:p></w:comment>
</w:comments>`,
      }),
    );
    expect(replyToComment(doc, "0", "I agree.", "Bob", "B")).toBe(true);
    expect(doc.comments.length).toBe(2);
    const reply = doc.comments.find((c) => c.author === "Bob")!;
    expect(reply.parentId).toBe("0");
    expect(reply.text).toBe("I agree.");
    // The reply is anchored to the same text as the parent.
    expect(doc.commentAnchors().get(reply.id)!.map((t) => t.text)).toEqual(["flagged"]);

    // Round-trip: threading survives save/load (via commentsExtended).
    const reloaded = DocxDocument.load(doc.save());
    const rr = reloaded.comments.find((c) => c.author === "Bob")!;
    expect(rr.parentId).toBe("0");
    expect(reloaded.pkg.text("word/commentsExtended.xml")).toContain("paraIdParent");

    // Deleting the parent removes the whole thread and all markers.
    expect(deleteComment(doc, "0")).toBe(true);
    expect(doc.comments.length).toBe(0);
    expect(doc.commentAnchors().size).toBe(0);
    const saved = DocxDocument.load(doc.save());
    expect(saved.pkg.text("word/document.xml")).not.toContain("commentReference");
  });
});

describe("tracked changes", () => {
  const XML = `<w:p>
    <w:r><w:t xml:space="preserve">kept </w:t></w:r>
    <w:ins w:id="1" w:author="A"><w:r><w:t xml:space="preserve">added </w:t></w:r></w:ins>
    <w:del w:id="2" w:author="A"><w:r><w:delText xml:space="preserve">removed </w:delText></w:r></w:del>
    <w:r><w:t>tail</w:t></w:r>
  </w:p>`;
  const textOf = (doc: DocxDocument): string => {
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    let out = "";
    for (const c of para.children) {
      const runs = c.type === "run" ? [c] : c.runs;
      for (const r of runs) for (const t of r.content) if (t.kind === "text") out += t.text;
    }
    return out;
  };

  it("final view keeps insertions and hides deletions", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML) }));
    expect(textOf(doc)).toBe("kept added tail");
  });

  it("markup view shows both, styled", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML) }));
    doc.setRevisionView("markup");
    expect(textOf(doc)).toBe("kept added removed tail");
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    const runs = para.children.filter((c) => c.type === "run");
    const added = runs.find((r) => r.type === "run" && r.content.some((c) => c.kind === "text" && c.text.includes("added")));
    const removed = runs.find((r) => r.type === "run" && r.content.some((c) => c.kind === "text" && c.text.includes("removed")));
    expect(added && added.type === "run" ? added.props.underline : "").toBe("single");
    expect(removed && removed.type === "run" ? removed.props.strike : false).toBe(true);
    // back to final
    doc.setRevisionView("final");
    expect(textOf(doc)).toBe("kept added tail");
  });
});

describe("OMML math", () => {
  const XML = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:r><w:t xml:space="preserve">Euler: </w:t></w:r>
    <m:oMath>
      <m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup><m:r><m:t>x</m:t></m:r></m:sup></m:sSup>
      <m:r><m:t>=1+</m:t></m:r>
      <m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
    </m:oMath>
  </w:p>`;

  it("parses equations into a math AST", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML) }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    const math = para.children.flatMap((c) => (c.type === "run" ? c.content : [])).find((c) => c.kind === "math");
    if (!math || math.kind !== "math") throw new Error("no math content");
    expect(math.nodes[0].t).toBe("sup");
    expect(math.nodes[1]).toEqual({ t: "run", text: "=1+" });
    expect(math.nodes[2].t).toBe("frac");
  });

  it("keeps normal math runs upright", () => {
    const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMath>
        <m:r><m:rPr><m:nor/></m:rPr><m:t>sin</m:t></m:r>
        <m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>cos</m:t></m:r>
        <m:r><m:t>x</m:t></m:r>
      </m:oMath>
    </w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    const math = para.children.flatMap((c) => (c.type === "run" ? c.content : [])).find((c) => c.kind === "math");
    if (!math || math.kind !== "math") throw new Error("no math content");
    expect(math.nodes).toEqual([
      { t: "run", text: "sincos", normal: true },
      { t: "run", text: "x" },
    ]);

    const texts = layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
      .filter((item) => item.kind === "text")
      .map((item) => item.text);
    expect(texts).toContain("sincos");
    expect(texts).toContain("𝑥");
  });

  it("lays math out 2D with Word's measured geometry", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML) }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const items = result.pages[0].items;
    const texts = items.filter((i) => i.kind === "text");
    const piece = (t: string) => {
      const it = texts.find((i) => i.kind === "text" && i.text === t);
      if (!it || it.kind !== "text") throw new Error("missing piece " + t);
      return it;
    };
    const base = piece("𝑒");
    const sup = piece("𝑥"); // first 𝑥 = the superscript
    // script raised by 4/11 of the base size, scaled to 8/11
    expect(base.baseline - sup.baseline).toBeCloseTo(base.font.size * (4 / 11), 1);
    expect(sup.font.size).toBeCloseTo(base.font.size * (8 / 11), 2);
    // fraction: numerator above, denominator below, rule between
    const den = piece("2");
    expect(den.baseline - base.baseline).toBeCloseTo(base.font.size * (5.5 / 11), 1);
    const rule = items.find((i) => i.kind === "rect" && i.fill === "#000000" && i.height < 2);
    expect(rule).toBeTruthy();
  });

  it("floors recursively nested math at scriptscript size without changing simple scripts", () => {
    const render = (script: string) => {
      const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:oMath><m:f>
          <m:num><m:r><m:t>N</m:t></m:r></m:num>
          <m:den><m:sSup>
            <m:e><m:r><m:t>B</m:t></m:r></m:e>
            <m:sup>${script}</m:sup>
          </m:sSup></m:den>
        </m:f></m:oMath>
      </w:p>`;
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
      return layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
        .filter((item) => item.kind === "text");
    };
    const piece = (items: ReturnType<typeof render>, text: string) => {
      const found = items.find((item) => item.kind === "text" && item.text === text);
      if (found?.kind !== "text") throw new Error(`missing ${text}`);
      return found;
    };

    const simple = render(`<m:r><m:t>S</m:t></m:r>`);
    const nested = render(
      `<m:f><m:num><m:r><m:t>X</m:t></m:r></m:num>` +
        `<m:den><m:r><m:t>Y</m:t></m:r></m:den></m:f>`,
    );
    const simpleBase = piece(simple, "𝐵");
    const simpleScript = piece(simple, "𝑆");
    const nestedNumerator = piece(nested, "𝑋");
    const nestedDenominator = piece(nested, "𝑌");

    // The simple p142-like denominator script keeps its calibrated 8/11
    // reduction from the surrounding 8/11 fraction part.
    expect(simpleScript.font.size).toBeCloseTo(simpleBase.font.size * (8 / 11), 5);
    // A fraction inside that script is already in scriptscript style. Its
    // children stay at the same floor instead of shrinking by 8/11 again.
    expect(nestedNumerator.font.size).toBeCloseTo(simpleScript.font.size, 5);
    expect(nestedDenominator.font.size).toBeCloseTo(simpleScript.font.size, 5);
  });

  it("spreads display integral limits without changing inline limits", () => {
    const integral =
      `<m:nary><m:naryPr><m:limLoc m:val="subSup"/></m:naryPr>` +
      `<m:sub><m:r><m:t>0</m:t></m:r></m:sub>` +
      `<m:sup><m:r><m:t>1</m:t></m:r></m:sup>` +
      `<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>`;
    const pieces = (display: boolean) => {
      const equation = display
        ? `<m:oMathPara><m:oMath>${integral}</m:oMath></m:oMathPara>`
        : `<m:oMath>${integral}</m:oMath>`;
      const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${equation}</w:p>`;
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
      const texts = layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
        .filter((item) => item.kind === "text");
      const piece = (text: string) => {
        const item = texts.find((candidate) => candidate.kind === "text" && candidate.text === text);
        if (item?.kind !== "text") throw new Error(`missing ${text}`);
        return item;
      };
      return { integral: piece("∫"), upper: piece("1"), lower: piece("0") };
    };

    const inline = pieces(false);
    const display = pieces(true);
    const inlineUpperGap = inline.integral.baseline - inline.upper.baseline;
    const inlineLowerGap = inline.lower.baseline - inline.integral.baseline;
    const displayUpperGap = display.integral.baseline - display.upper.baseline;
    const displayLowerGap = display.lower.baseline - display.integral.baseline;

    // Dense p7 Word control: display limits extend 5.36pt farther upward and
    // 7.37pt farther downward than the current, already-accurate inline stack.
    expect(inlineUpperGap).toBeCloseTo(inline.integral.font.size * (7.25 / 11), 3);
    expect(inlineLowerGap).toBeCloseTo(inline.integral.font.size * (4 / 11), 3);
    expect(displayUpperGap - inlineUpperGap).toBeCloseTo(
      inline.integral.font.size * (5.359 / 12),
      3,
    );
    expect(displayLowerGap - inlineLowerGap).toBeCloseTo(
      inline.integral.font.size * (7.369 / 12),
      3,
    );
  });

  it("marks m:oMathPara as a display equation and honors noBar", () => {
    const XML_DISPLAY = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMathPara><m:oMath>
        <m:f><m:fPr><m:type m:val="noBar"/></m:fPr>
          <m:num><m:r><m:t>n</m:t></m:r></m:num><m:den><m:r><m:t>k</m:t></m:r></m:den></m:f>
      </m:oMath></m:oMathPara></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML_DISPLAY) }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    const math = para.children.flatMap((c) => (c.type === "run" ? c.content : [])).find((c) => c.kind === "math");
    if (!math || math.kind !== "math") throw new Error("no math content");
    expect(math.display).toBe(true);
    const frac = math.nodes[0];
    if (frac.t !== "frac") throw new Error("expected frac");
    expect(frac.bar).toBe(false);
  });

  it("renders a linear fraction inline while keeping a default fraction stacked", () => {
    const render = (type = "") => {
      const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:oMath><m:f>${type}<m:num><m:r><m:t>1</m:t></m:r></m:num>` +
        `<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>
      </w:p>`;
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
      const para = doc.sections[0].blocks[0];
      if (para.type !== "paragraph") throw new Error();
      const math = para.children.flatMap((c) => (c.type === "run" ? c.content : []))
        .find((content) => content.kind === "math");
      if (!math || math.kind !== "math") throw new Error("no math content");
      return {
        math,
        items: layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items,
      };
    };

    const linear = render(`<m:fPr><m:type m:val="lin"/></m:fPr>`);
    expect(linear.math.nodes).toEqual([
      { t: "run", text: "1" },
      { t: "run", text: "/", normal: true },
      { t: "run", text: "2" },
    ]);
    const linearNumbers = linear.items.filter(
      (item) => item.kind === "text" && (item.text === "1" || item.text === "2"),
    );
    expect(linearNumbers[0]?.kind).toBe("text");
    expect(linearNumbers[1]?.kind).toBe("text");
    if (linearNumbers[0]?.kind !== "text" || linearNumbers[1]?.kind !== "text") return;
    expect(linearNumbers[0].baseline).toBeCloseTo(linearNumbers[1].baseline, 3);

    const stacked = render();
    expect(stacked.math.nodes[0]?.t).toBe("frac");
    const stackedRule = stacked.items.find(
      (item) => item.kind === "rect" && item.fill === "#000000" && item.height < 2,
    );
    expect(stackedRule).toBeDefined();
  });

  it("adds text leading to tall inline math under double spacing", () => {
    const render = (equation: string, multiple: 1 | 2) => {
      const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:pPr><w:spacing w:after="0" w:line="${multiple * 240}" w:lineRule="auto"/></w:pPr>
        <m:oMath>${equation}</m:oMath><w:r><w:t> equation</w:t></w:r>
      </w:p><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
        <w:r><w:t>NEXT</w:t></w:r></w:p>`;
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
      const next = layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
        .find((item) => item.kind === "text" && item.text === "NEXT");
      if (next?.kind !== "text") throw new Error("next paragraph missing");
      return next.lineTop;
    };
    const simple = `<m:r><m:t>x</m:t></m:r>`;
    const inner = `<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num>` +
      `<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>`;
    const tall = `<m:f><m:num>${inner}</m:num>` +
      `<m:den><m:r><m:t>3</m:t></m:r></m:den></m:f>`;

    const simpleLead = render(simple, 2) - render(simple, 1);
    const tallLead = render(tall, 2) - render(tall, 1);
    expect(tallLead).toBeCloseTo(simpleLead, 3);
  });

  it("preserves explicit line breaks inside OMML runs", () => {
    const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMath>
        <m:r><m:t>A</m:t></m:r>
        <m:r><w:br/></m:r>
        <m:r><m:t>B</m:t></m:r>
      </m:oMath>
    </w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    const run = para.children[0];
    if (run.type !== "run") throw new Error();
    expect(run.content.map((content) => content.kind)).toEqual(["math", "break", "math"]);

    const texts = layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
      .filter((item) => item.kind === "text");
    const a = texts.find((item) => item.kind === "text" && item.text === "𝐴");
    const b = texts.find((item) => item.kind === "text" && item.text === "𝐵");
    expect(a?.kind).toBe("text");
    expect(b?.kind).toBe("text");
    if (a?.kind !== "text" || b?.kind !== "text") return;
    expect(b.lineTop).toBeGreaterThan(a.lineTop);
  });

  it("wraps display math at an exposed operator but keeps nested math atomic", () => {
    const a = "A".repeat(36);
    const b = "B".repeat(36);
    const render = (equation: string) => {
      const xml = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:oMathPara><m:oMath>${equation}</m:oMath></m:oMathPara>
      </w:p>`;
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(xml) }));
      return layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items
        .filter((item) => item.kind === "text");
    };
    const pieces = (items: ReturnType<typeof render>) => {
      const left = items.find((item) => item.kind === "text" && item.text.includes("𝐴"));
      const plus = items.find((item) => item.kind === "text" && item.text === "+");
      const right = items.find((item) => item.kind === "text" && item.text.includes("𝐵"));
      if (left?.kind !== "text" || plus?.kind !== "text" || right?.kind !== "text") {
        throw new Error("missing equation pieces");
      }
      return { left, plus, right };
    };

    const exposed = pieces(render(
      `<m:d><m:e><m:r><m:t>${a}</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>${b}</m:t></m:r></m:e></m:d>`,
    ));
    expect(exposed.right.lineTop).toBeGreaterThan(exposed.left.lineTop);
    expect(exposed.plus.lineTop).toBe(exposed.right.lineTop);

    const nested = pieces(render(
      `<m:f><m:num><m:r><m:t>${a}</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>${b}</m:t></m:r></m:num>` +
        `<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>`,
    ));
    expect(nested.right.lineTop).toBe(nested.left.lineTop);
  });

  it("display fractions use full-size num/den and center the equation", () => {
    const XML_DISPLAY = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMathPara><m:oMath>
        <m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
      </m:oMath></m:oMathPara></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML_DISPLAY) }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const texts = result.pages[0].items.filter((i) => i.kind === "text");
    const numX = texts.find((i) => i.kind === "text" && i.text === "𝑥");
    const denN = texts.find((i) => i.kind === "text" && i.text === "2");
    if (!numX || numX.kind !== "text" || !denN || denN.kind !== "text") throw new Error("missing pieces");
    // Full base size (not the 8/11 inline-fraction script scale).
    expect(numX.font.size).toBeCloseTo(denN.font.size, 5);
    expect(numX.font.size).toBeGreaterThan(14); // ~11pt in px
    // Centered on the content column: left edge well past the left margin.
    expect(numX.x).toBeGreaterThan(150);
  });
});
