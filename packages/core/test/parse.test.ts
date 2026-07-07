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
  it("extracts a legible linear string from equations", () => {
    const XML = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:r><w:t xml:space="preserve">Euler: </w:t></w:r>
      <m:oMath>
        <m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup><m:r><m:t>x</m:t></m:r></m:sup></m:sSup>
        <m:r><m:t>=1+</m:t></m:r>
        <m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>1!</m:t></m:r></m:den></m:f>
      </m:oMath>
    </w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(XML) }));
    const para = doc.sections[0].blocks[0];
    if (para.type !== "paragraph") throw new Error();
    let text = "";
    for (const c of para.children) {
      const runs = c.type === "run" ? [c] : c.runs;
      for (const r of runs) for (const t of r.content) if (t.kind === "text") text += t.text;
    }
    expect(text).toBe("Euler: e^x=1+x⁄" + "1!");
  });
});
