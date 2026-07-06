import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { parseXml, child, attr } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

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
});
