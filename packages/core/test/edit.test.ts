import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { applyRunFormat, SelectionSegment, selectionTextLogical } from "../src/edit/commands.js";
import { addComment } from "../src/edit/comments.js";
import { setListType, setListLevel } from "../src/edit/lists.js";
import { setLink, removeLink, linkAt } from "../src/edit/links.js";
import { adjustIndent, paragraphDividerAt, setDropCapAt, setParagraphDivider, setParagraphSpacing } from "../src/edit/paragraph.js";
import { findAll, replaceAll, transformCase } from "../src/edit/find.js";
import { applyTableOp, cellShadingAt, resizeDrawing } from "../src/edit/tables.js";
import {
  drawingRotation,
  imageAltText,
  replaceImageBlip,
  setDrawingOrder,
  setDrawingRotation,
  setFloatingPagePosition,
  setImageAltText,
} from "../src/edit/images.js";
import { insertFootnote } from "../src/edit/notes.js";
import { insertDateTimeField, insertField, insertPageField } from "../src/edit/fields.js";
import { insertBlankPageAt, insertBreakAt, insertCoverPage, sectionContextAt } from "../src/edit/sections.js";
import { drawingLineStyle, drawingWordArtText, insertInkAt, insertShapeAt, insertWordArtAt, isDrawingWordArt, setDrawingLineStyle, setDrawingWordArtText, type ShapePreset, type WordArtPreset } from "../src/edit/drawings.js";
import { buildChartXml, insertChartAt, setChartData } from "../src/edit/charts.js";
import { buildSmartArtDataXml, buildSmartArtDrawingXml, insertSmartArtAt, setSmartArtData, setSmartArtFill, setSmartArtNodeText, setSmartArtTextFormat, smartArtFillColor, smartArtTextFormat } from "../src/edit/smartart.js";
import { insertEmbeddedObjectAt, insertModel3DAt, insertWebVideoAt } from "../src/edit/objects.js";
import { buildOlePackage, extractOlePackage } from "../src/parse/ole.js";
import { insertBookmarkAroundSelection, insertBookmarkAt, insertCrossReference, listBookmarks, validBookmarkName } from "../src/edit/references.js";
import { linearizeMath, parseMathLinear, setMathLinear, insertMathAt, mathLinearOf, isLinearSafe } from "../src/edit/math.js";
import { XmlElement, localName } from "../src/xml.js";
import { serializeXml, parseXml } from "../src/xml.js";
import { makeDocx, makeDocxWithMedia, wrapDocument, p, W_NS } from "./helpers.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { layoutDocument } from "../src/layout/engine.js";
import type { ChartData, SmartArtData } from "../src/model.js";
import { Package } from "../src/zip.js";
import * as CFB from "cfb";

function loadDoc(body: string, extra: Record<string, string> = {}) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...extra }));
}

function firstRun(doc: DocxDocument, blockIdx = 0): { para: Paragraph; run: Run } {
  const para = doc.sections[0].blocks[blockIdx] as Paragraph;
  const run = para.children[0] as Run;
  return { para, run };
}

function textOf(para: Paragraph): string {
  let out = "";
  for (const c of para.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) for (const rc of r.content) if (rc.kind === "text") out += rc.text;
  }
  return out;
}

function segFor(run: Run, start: number, end: number): SelectionSegment {
  const t = (run.content.find((c) => c.kind === "text") as TextContent | undefined)?.srcT ?? null;
  return { run, t: t as SelectionSegment["t"], start, end, props: run.props };
}

describe("package content types", () => {
  const customProperties = `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`;
  const contentTypes = (customOverride: string, prefix = "") => `<?xml version="1.0"?>
<${prefix}Types xmlns${prefix ? `:${prefix.slice(0, -1)}` : ""}="http://schemas.openxmlformats.org/package/2006/content-types">
  <${prefix}Default Extension="xml" ContentType="application/xml"/>
  <${prefix}Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <${prefix}Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${customOverride}
</${prefix}Types>`;
  const customContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";

  it("adds the custom-properties content type when the preserved part is missing its override", () => {
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(p("Custom properties")),
      "docProps/custom.xml": customProperties,
      "[Content_Types].xml": contentTypes("", "ct:"),
    }));

    const savedContentTypes = strFromU8(unzipSync(doc.save())["[Content_Types].xml"]);
    expect(savedContentTypes).toContain("<ct:Override");
    expect(savedContentTypes).toContain(`PartName="/docProps/custom.xml" ContentType="${customContentType}"`);
  });

  it("preserves an existing custom-properties content type without duplicating it", () => {
    const override = `<Override PartName="/docProps/custom.xml" ContentType="${customContentType}"/>`;
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(p("Custom properties")),
      "docProps/custom.xml": customProperties,
      "[Content_Types].xml": contentTypes(override),
    }));

    const savedContentTypes = strFromU8(unzipSync(doc.save())["[Content_Types].xml"]);
    expect(savedContentTypes.match(/PartName="\/docProps\/custom\.xml"/g)).toHaveLength(1);
    expect(savedContentTypes).toContain(`PartName="/docProps/custom.xml" ContentType="${customContentType}"`);
  });
});

describe("no-op package round trip", () => {
  it("preserves original modeled XML bytes until a retained tree changes", () => {
    const crlf = (xml: string) => xml.replace(/\n/g, "\r\n");
    const documentXml = crlf(wrapDocument(p("Original")));
    const relationships = crlf(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`);
    const contentTypes = crlf(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`);
    const headerXml = crlf(`<?xml version="1.0"?>
<w:hdr ${W_NS}>
  ${p("Header")}
</w:hdr>`);
    const footerXml = crlf(`<?xml version="1.0"?>
<w:ftr ${W_NS}>
  ${p("Footer")}
</w:ftr>`);
    const originalParts = {
      "word/document.xml": documentXml,
      "word/_rels/document.xml.rels": relationships,
      "[Content_Types].xml": contentTypes,
      "word/header1.xml": headerXml,
      "word/footer1.xml": footerXml,
    };
    const doc = DocxDocument.load(makeDocx(originalParts));

    const untouched = unzipSync(doc.save());
    for (const [part, xml] of Object.entries(originalParts)) {
      expect(untouched[part]).toEqual(strToU8(xml));
    }

    const { run } = firstRun(doc);
    const text = run.content.find((item) => item.kind === "text") as TextContent;
    text.srcT!.text = "Edited";
    const edited = unzipSync(doc.save());
    expect(edited["word/document.xml"]).not.toEqual(strToU8(documentXml));
    expect(strFromU8(edited["word/document.xml"])).toContain(">Edited</w:t>");
    for (const part of ["word/_rels/document.xml.rels", "[Content_Types].xml", "word/header1.xml", "word/footer1.xml"]) {
      expect(edited[part]).toEqual(strToU8(originalParts[part as keyof typeof originalParts]));
    }
    expect(textOf(firstRun(DocxDocument.load(doc.save())).para)).toBe("Edited");

    const header = doc.headers.get("rId1")?.blocks[0];
    if (!header || header.type !== "paragraph") throw new Error("header paragraph missing");
    const headerRun = header.children[0];
    if (!headerRun || headerRun.type !== "run") throw new Error("header run missing");
    const headerText = headerRun.content.find((item) => item.kind === "text") as TextContent;
    headerText.srcT!.text = "Edited header";
    const headerEdited = unzipSync(doc.save());
    expect(strFromU8(headerEdited["word/header1.xml"])).toContain(">Edited header</w:t>");
    expect(headerEdited["word/footer1.xml"]).toEqual(strToU8(footerXml));
    const reloadedHeader = DocxDocument.load(doc.save()).headers.get("rId1")?.blocks[0];
    if (!reloadedHeader || reloadedHeader.type !== "paragraph") throw new Error("reloaded header missing");
    expect(textOf(reloadedHeader)).toBe("Edited header");
  });

  it("canonicalizes placeholder percentage grids for Google Docs without changing Word autofit", () => {
    const doc = loadDoc(
      `<w:tbl><w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>Key</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>Description of the item</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>A much longer description cell that should dominate the width</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
      `</w:sectPr>`,
    );

    const saved = strFromU8(unzipSync(doc.save())["word/document.xml"]);
    expect(saved).toContain('<w:tblW w:type="pct" w:w="5000"/>');
    expect([...saved.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((match) => Number(match[1])))
      .toEqual([600, 900, 7526]);
  });

  it("keeps an empty-paragraph caret anchor in memory without saving its synthetic run", () => {
    const doc = loadDoc(
      `<w:p><w:pPr><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr></w:p>` +
      p("Following paragraph"),
    );
    const paragraph = doc.sections[0].blocks[0] as Paragraph;
    const anchor = paragraph.children[0] as Run;
    expect(anchor.content[0]).toMatchObject({ kind: "text", text: "" });
    expect((anchor.content[0] as TextContent).srcT).toBeDefined();

    const savedDocument = parseXml(strFromU8(unzipSync(doc.save())["word/document.xml"]));
    const body = savedDocument.children.find((item) => localName(item.name) === "body")!;
    const savedParagraph = body.children.find((item) => localName(item.name) === "p")!;
    expect(savedParagraph.children.map((item) => localName(item.name))).toEqual(["pPr"]);

    (anchor.content[0] as TextContent).srcT!.text = "Typed";
    expect(strFromU8(unzipSync(doc.save())["word/document.xml"])).toContain(">Typed</w:t>");
  });

  it("does not add a fontTable relationship when the preserved part had none", () => {
    const relationships = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
    const doc = loadDoc(p("Font table"), {
      "word/fontTable.xml": `<w:fonts ${W_NS}/>`,
      "word/_rels/document.xml.rels": relationships,
    });

    const saved = unzipSync(doc.save());
    expect(saved["word/fontTable.xml"]).toBeDefined();
    expect(strFromU8(saved["word/_rels/document.xml.rels"])).not.toContain("/fontTable");
  });
});

describe("run formatting commands", () => {
  it("formats a full run in place (no split)", () => {
    const doc = loadDoc(p("Hello world"));
    const { run } = firstRun(doc);
    applyRunFormat(doc, [segFor(run, 0, 11)], { bold: true });
    const { para: after } = firstRun(doc);
    expect(after.children.length).toBe(1);
    const r = after.children[0] as Run;
    expect(r.props.bold).toBe(true);
    expect(textOf(after)).toBe("Hello world");
  });

  it("splits a run for a partial selection", () => {
    const doc = loadDoc(p("Hello brave world"));
    const { run } = firstRun(doc);
    // Select "brave" (chars 6..11)
    applyRunFormat(doc, [segFor(run, 6, 11)], { bold: true, color: "#FF0000" });
    const { para: after } = firstRun(doc);
    expect(textOf(after)).toBe("Hello brave world");
    const runs = after.children.filter((c) => c.type === "run") as Run[];
    expect(runs.length).toBe(3);
    expect(runs[0].props.bold).toBeUndefined();
    expect(runs[1].props.bold).toBe(true);
    expect(runs[1].props.color).toBe("#FF0000");
    expect((runs[1].content[0] as TextContent).text).toBe("brave");
    expect(runs[2].props.bold).toBeUndefined();
  });

  it("preserves existing run formatting on split fragments", () => {
    const doc = loadDoc(
      `<w:p><w:r><w:rPr><w:i/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">italic text here</w:t></w:r></w:p>`,
    );
    const { run } = firstRun(doc);
    applyRunFormat(doc, [segFor(run, 7, 11)], { bold: true });
    const runs = (doc.sections[0].blocks[0] as Paragraph).children as Run[];
    expect(runs.length).toBe(3);
    for (const r of runs) {
      expect(r.props.italic).toBe(true); // inherited via cloned rPr
      expect(r.props.size).toBeCloseTo((14 * 4) / 3, 2);
    }
    expect(runs[1].props.bold).toBe(true);
  });

  it("survives a save/load round trip with edits", () => {
    const doc = loadDoc(p("Hello brave world"));
    const { run } = firstRun(doc);
    applyRunFormat(doc, [segFor(run, 6, 11)], { bold: true, fontSizePt: 18, highlight: "yellow" });
    const bytes = doc.save();
    const reloaded = DocxDocument.load(bytes);
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    expect(textOf(para)).toBe("Hello brave world");
    const runs = para.children.filter((c) => c.type === "run") as Run[];
    expect(runs.length).toBe(3);
    expect(runs[1].props.bold).toBe(true);
    expect(runs[1].props.size).toBeCloseTo(24, 1); // 18pt = 24px
    expect(runs[1].props.highlight).toBe("#ffff00");
  });

  it("preserves Symbol-font source bytes when formatting decoded text", () => {
    const source = "\uF067\uF020\uF071\uF02E";
    const doc = loadDoc(
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t>${source}</w:t></w:r></w:p>`,
    );
    const { run } = firstRun(doc);
    expect(textOf(firstRun(doc).para)).toBe("γ θ.");

    applyRunFormat(doc, [segFor(run, 0, 1)], { bold: true });

    const saved = DocxDocument.load(doc.save());
    const xml = saved.pkg.text("word/document.xml");
    expect(xml).toContain(source.slice(0, 1));
    expect(xml).toContain(source.slice(1));
    expect(xml).not.toContain("γ θ.");
    expect(textOf(firstRun(saved).para)).toBe("γ θ.");
  });

  it("keeps unrelated parts byte-identical on save", () => {
    const styles = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="X"><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;
    const doc = loadDoc(p("text"), { "word/styles.xml": styles });
    const { run } = firstRun(doc);
    applyRunFormat(doc, [segFor(run, 0, 4)], { italic: true });
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.pkg.text("word/styles.xml")).toBe(styles);
  });

  it("serializes entities safely", () => {
    const root = parseXml(`<w:p w:val="a&amp;b"><w:t>x &lt; y &amp; z</w:t></w:p>`);
    const out = serializeXml(root);
    expect(out).toBe(`<w:p w:val="a&amp;b"><w:t>x &lt; y &amp; z</w:t></w:p>`);
  });

  it("removes properties when patch value is null", () => {
    const doc = loadDoc(
      `<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>lit</w:t></w:r></w:p>`,
    );
    const { run } = firstRun(doc);
    expect(run.props.highlight).toBe("#ffff00");
    applyRunFormat(doc, [segFor(run, 0, 3)], { highlight: null });
    const after = firstRun(doc).run;
    expect(after.props.highlight).toBeUndefined();
  });
});

describe("paragraph styles", () => {
  it("applying an undeclared built-in heading injects Word's definition", async () => {
    const { setParagraphStyle, paragraphStyleIdOf } = await import("../src/edit/blocks.js");
    const styles = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`;
    const doc = loadDoc(p("make me a heading"), { "word/styles.xml": styles });
    expect(doc.styles.byId.has("Heading1")).toBe(false);
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    expect(setParagraphStyle(doc, [t as never], "Heading1")).toBe(true);
    expect(paragraphStyleIdOf(doc, t as never)).toBe("Heading1");
    // The injected definition resolves: heading renders larger and colored.
    expect(doc.styles.byId.get("Heading1")?.name).toBe("Heading 1");
    const para = firstRun(doc).para;
    const props = doc.effectiveRunProps(para, (para.children[0] as Run).props);
    expect(props.size).toBeCloseTo((16 * 4) / 3, 1); // 32 half-points = 16pt
    expect(props.color?.toLowerCase()).toBe("#2f5496");
    // Round-trips: the reloaded file still declares and resolves the style.
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.styles.byId.get("Heading1")?.name).toBe("Heading 1");
  });
});

describe("undo/redo history", () => {
  it("retains text identities and model generation when the editor accepts a text-only undo", async () => {
    const { EditHistory } = await import("../src/edit/history.js");
    const doc = loadDoc(p("Hello world"));
    const history = new EditHistory(doc);
    const { run } = firstRun(doc);
    const content = run.content[0] as TextContent;
    const t = content.srcT!;
    const modelVersion = doc.modelVersion;
    history.applyTextChanges = (changes) => {
      if (changes.length !== 1 || changes[0] !== t) return false;
      content.text = t.text;
      return true;
    };

    history.checkpoint("typing");
    t.text = "Hello brave world";
    content.text = t.text;
    expect(history.undo()).toBe(true);
    expect(history.lastTextChanges).toEqual([t]);
    expect(t.text).toBe("Hello world");
    expect(content.text).toBe("Hello world");
    expect(doc.modelVersion).toBe(modelVersion);

    expect(history.redo()).toBe(true);
    expect(history.lastTextChanges).toEqual([t]);
    expect(t.text).toBe("Hello brave world");
    expect(content.text).toBe("Hello brave world");
    expect(doc.modelVersion).toBe(modelVersion);
  });

  it("undoes and redoes a text mutation", async () => {
    const { EditHistory } = await import("../src/edit/history.js");
    const doc = loadDoc(p("Hello world"));
    const history = new EditHistory(doc);
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT as { text: string };

    history.checkpoint("typing");
    t.text = "Hello brave world";
    doc.refresh();
    expect(textOf(firstRun(doc).para)).toBe("Hello brave world");

    expect(history.undo()).toBe(true);
    expect(textOf(firstRun(doc).para)).toBe("Hello world");
    expect(history.redo()).toBe(true);
    expect(textOf(firstRun(doc).para)).toBe("Hello brave world");
  });

  it("coalesces rapid same-kind checkpoints into one undo step", async () => {
    const { EditHistory } = await import("../src/edit/history.js");
    const doc = loadDoc(p("ab"));
    const history = new EditHistory(doc);
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT as { text: string };

    history.checkpoint("typing");
    t.text = "abc";
    history.checkpoint("typing"); // coalesced
    t.text = "abcd";
    doc.refresh();

    expect(history.undo()).toBe(true);
    expect(textOf(firstRun(doc).para)).toBe("ab"); // whole burst undone
    expect(history.undo()).toBe(false);
  });

  it("formatting commands are undoable", async () => {
    const { EditHistory } = await import("../src/edit/history.js");
    const doc = loadDoc(p("Hello brave world"));
    const history = new EditHistory(doc);
    const { run } = firstRun(doc);
    history.checkpoint();
    const modelVersion = doc.modelVersion;
    applyRunFormat(doc, [segFor(run, 6, 11)], { bold: true });
    expect((firstRun(doc).para.children as Run[]).length).toBe(3);
    history.undo();
    expect(history.lastTextChanges).toBeNull();
    expect(doc.modelVersion).toBeGreaterThan(modelVersion);
    expect((firstRun(doc).para.children as Run[]).length).toBe(1);
    expect(textOf(firstRun(doc).para)).toBe("Hello brave world");
  });
});

describe("local paragraph split reparsing", () => {
  it("splices two parsed paragraphs while preserving model generation and unchanged blocks", () => {
    const doc = loadDoc(p("before") + p("alpha beta") + p("after"));
    const first = doc.sections[0].blocks[0];
    const original = doc.sections[0].blocks[1] as Paragraph;
    const last = doc.sections[0].blocks[2];
    const originalRun = original.children[0] as Run;
    const originalText = (originalRun.content[0] as TextContent).srcT!;
    const source = original.src!;
    const body = doc.docRoot.children.find((element) => localName(element.name) === "body")!;
    const sourceIndex = body.children.indexOf(source);
    const afterText: XmlElement = {
      name: originalText.name,
      attrs: { ...originalText.attrs, "xml:space": "preserve" },
      children: [],
      text: " beta",
    };
    const afterSource: XmlElement = {
      name: source.name,
      attrs: {},
      text: "",
      children: [{ name: originalRun.src!.name, attrs: {}, text: "", children: [afterText] }],
    };
    originalText.text = "alpha";
    body.children.splice(sourceIndex + 1, 0, afterSource);

    const version = doc.modelVersion;
    const reparsed = doc.reparseDirectBodyParagraphSplit(source, afterSource);
    expect(reparsed).not.toBeNull();
    expect(doc.modelVersion).toBe(version);
    expect(doc.sections[0].blocks).toHaveLength(4);
    expect(doc.sections[0].blocks[0]).toBe(first);
    expect(doc.sections[0].blocks[3]).toBe(last);
    expect(textOf(reparsed!.before)).toBe("alpha");
    expect(textOf(reparsed!.after)).toBe(" beta");
    const afterRun = reparsed!.after.children[0] as Run;
    expect((afterRun.content[0] as TextContent).srcT).toBe(afterText);

    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.sections[0].blocks.map((block) => block.type === "paragraph" ? textOf(block) : "table"))
      .toEqual(["before", "alpha", " beta", "after"]);
  });

  it.each([
    ["bookmark", `<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>alpha beta</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`],
    ["revision", `<w:p><w:ins w:id="1"><w:r><w:t>alpha beta</w:t></w:r></w:ins></w:p>`],
    ["content control", `<w:p><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>alpha beta</w:t></w:r></w:sdtContent></w:sdt></w:p>`],
    ["field", `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:t>alpha beta</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`],
    ["section break", `<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>alpha beta</w:t></w:r></w:p>`],
  ])("falls back to a full refresh for a %s paragraph", (_name, paragraph) => {
    const doc = loadDoc(paragraph + p("after"));
    const source = doc.sections[0].blocks[0].src!;
    const body = doc.docRoot.children.find((element) => localName(element.name) === "body")!;
    const sourceIndex = body.children.indexOf(source);
    const afterSource = parseXml(`<w:p><w:r><w:t>tail</w:t></w:r></w:p>`);
    body.children.splice(sourceIndex + 1, 0, afterSource);

    expect(doc.reparseDirectBodyParagraphSplit(source, afterSource)).toBeNull();
    const version = doc.modelVersion;
    doc.refresh();
    expect(doc.modelVersion).toBeGreaterThan(version);
  });
});

describe("block commands", () => {
  it("inserts a table after the caret paragraph and round-trips", async () => {
    const { insertTableAfter } = await import("../src/edit/blocks.js");
    const doc = loadDoc(p("before") + p("after"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    expect(insertTableAfter(doc, t as never, 2, 3)).toBe(true);
    const blocks = doc.sections[0].blocks;
    expect(blocks[1].type).toBe("table");
    if (blocks[1].type !== "table") return;
    expect(blocks[1].rows.length).toBe(2);
    expect(blocks[1].rows[0].cells.length).toBe(3);
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.sections[0].blocks[1].type).toBe("table");
  });

  it("sets paragraph alignment", async () => {
    const { setParagraphAlignment } = await import("../src/edit/blocks.js");
    const doc = loadDoc(p("some text"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    setParagraphAlignment(doc, [t as never], "center");
    expect(firstRun(doc).para.props.alignment).toBe("center");
  });

  it("changes margins and orientation", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      p("text") + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    setPageLayout(doc, { margins: { left: 0.5, right: 0.5 }, orientation: "landscape" });
    const sp = doc.sections[0].props;
    expect(sp.marginLeft).toBeCloseTo(48, 1); // 0.5in = 48px
    expect(sp.pageWidth).toBeGreaterThan(sp.pageHeight); // landscape
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.sections[0].props.pageWidth).toBeGreaterThan(reloaded.sections[0].props.pageHeight);
  });

  it("saves and reopens a custom page border color", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      p("text") + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    expect(setPageLayout(doc, { pageBorders: { sz: 12, color: "#C62828" } })).toBe(true);
    const saved = strFromU8(unzipSync(doc.save())["word/document.xml"]);
    expect(saved.match(/w:color="C62828"/g)).toHaveLength(4);
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.sections[0].props.pageBorders?.top).toMatchObject({ color: "#C62828", width: 2 });
    expect(reloaded.sections[0].props.pageBorders?.right).toMatchObject({ color: "#C62828", width: 2 });
  });

  it("saves and reopens Word's native line-between-columns setting", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      p("left") + p("right") +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    expect(setPageLayout(doc, { columns: 2, columnSeparator: true })).toBe(true);
    expect(doc.sections[0].props.columns).toMatchObject({ count: 2, sep: true });
    const saved = strFromU8(unzipSync(doc.save())["word/document.xml"]);
    expect(saved).toMatch(/<w:cols\b[^>]*w:num="2"[^>]*w:sep="1"/);
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.sections[0].props.columns).toMatchObject({ count: 2, sep: true });
  });

  it("saves and reopens mirrored margins through settings.xml", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      p("text") + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
      {
        "word/settings.xml": `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:proofState w:spelling="clean"/><w:compat/></w:settings>`,
      },
    );
    setPageLayout(doc, {
      margins: { top: 1, right: 1, bottom: 1, left: 1.25 },
      mirrorMargins: true,
    });
    expect(doc.mirrorMargins).toBe(true);
    expect(doc.sections[0].props.marginLeft).toBeCloseTo(120, 1);

    const saved = unzipSync(doc.save());
    const settings = strFromU8(saved["word/settings.xml"]);
    const rels = strFromU8(saved["word/_rels/document.xml.rels"]);
    const contentTypes = strFromU8(saved["[Content_Types].xml"]);
    expect(settings).toContain("<w:mirrorMargins");
    expect(settings.indexOf("<w:zoom")).toBeLessThan(settings.indexOf("<w:mirrorMargins"));
    expect(settings.indexOf("<w:mirrorMargins")).toBeLessThan(settings.indexOf("<w:proofState"));
    expect(settings.indexOf("<w:mirrorMargins")).toBeLessThan(settings.indexOf("<w:compat"));
    expect(rels).toContain("/relationships/settings");
    expect(contentTypes).toContain("wordprocessingml.settings+xml");

    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.mirrorMargins).toBe(true);
    expect(reloaded.sections[0].props.marginLeft).toBeCloseTo(120, 1);
    setPageLayout(reloaded, { mirrorMargins: false });
    expect(reloaded.mirrorMargins).toBe(false);
    expect(strFromU8(unzipSync(reloaded.save())["word/settings.xml"])).not.toContain("mirrorMargins");
  });

  it("round-trips every supported page dimension", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const sizes = [
      [8.5, 11], [8.5, 14], [3.5, 5], [4, 6], [5, 7], [8, 10],
      [8.27, 11.69], [4.13, 5.83], [4.13, 9.5],
    ] as const;
    for (const [width, height] of sizes) {
      const doc = loadDoc(
        p("text") + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
      );
      setPageLayout(doc, { size: { width, height } });
      const saved = doc.save();
      const documentXml = strFromU8(unzipSync(saved)["word/document.xml"]);
      expect(documentXml).toContain(
        `<w:pgSz w:w="${Math.round(width * 1440)}" w:h="${Math.round(height * 1440)}"`,
      );
      const reloaded = DocxDocument.load(saved);
      expect(reloaded.sections[0].props.pageWidth).toBeCloseTo(width * 96, 0);
      expect(reloaded.sections[0].props.pageHeight).toBeCloseTo(height * 96, 0);
    }
  });

  it("scopes section margins while mirror mode remains document-global", async () => {
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const doc = loadDoc(
      `<w:p><w:pPr>${section}</w:pPr><w:r><w:t>first</w:t></w:r></w:p>` + p("second") + section,
    );
    const sectPrs: XmlElement[] = [];
    const collect = (element: XmlElement) => {
      if (localName(element.name) === "sectPr") sectPrs.push(element);
      for (const child of element.children) collect(child);
    };
    collect(doc.docRoot);
    expect(sectPrs).toHaveLength(2);

    setPageLayout(doc, { margins: { left: 0.5 }, mirrorMargins: true }, sectPrs[0]);
    expect(doc.sections[0].props.marginLeft).toBeCloseTo(48, 1);
    expect(doc.sections[1].props.marginLeft).toBeCloseTo(96, 1);
    expect(doc.mirrorMargins).toBe(true);

    setPageLayout(doc, { margins: { left: 0.75 }, mirrorMargins: false });
    expect(doc.sections[0].props.marginLeft).toBeCloseTo(72, 1);
    expect(doc.sections[1].props.marginLeft).toBeCloseTo(72, 1);
    expect(doc.mirrorMargins).toBe(false);
  });
});

describe("table operations", () => {
  const tableDoc = () =>
    loadDoc(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
       <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
       <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
       </w:tbl>` + p("after"),
    );
  const caretIn = (doc: DocxDocument, text: string) => {
    let found: unknown = null;
    const walk = (el: { children: { name: string; text: string; children: unknown[] }[] } & { name?: string; text?: string }) => {
      for (const c of el.children as never[]) {
        const e = c as { name: string; text: string; children: never[] };
        if (e.name.endsWith("t") && e.text === text) found = e;
        walk(e);
      }
    };
    walk(doc.editableRoots()[0] as never);
    return found as never;
  };

  it("drag-resizing a column stamps dxa tcW on every cell, spans included", async () => {
    const { resizeTableColumn } = await import("../src/edit/tables.js");
    // Percent-width table with a gridSpan row and NO tcW anywhere — the shape
    // the layout treats as autofit (untrusted grid). A drag must convert it
    // to a fully-declared fixed grid or the new widths are ignored.
    const doc = loadDoc(
      `<w:tbl><w:tblPr><w:tblW w:type="pct" w:w="5000"/></w:tblPr>
       <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
       <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc></w:tr>
       <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>AB2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc></w:tr>
       </w:tbl>` + p("after"),
    );
    const tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table" || !tbl.src) throw new Error("not a table");
    expect(resizeTableColumn(doc, tbl.src, 1, 30, [200, 200, 200])).toBe(true);
    const xml = serializeXml(tbl.src);
    // Grid rewritten from the rendered widths with the delta applied.
    const gridW = [...xml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => parseInt(m[1], 10));
    expect(gridW.length).toBe(3);
    expect(gridW[0]).toBeGreaterThan(gridW[1]); // +30px went to col 1, out of col 2
    // Every cell now declares a fixed width — spanned cell covers two columns.
    const tcW = [...xml.matchAll(/<w:tcW w:type="dxa" w:w="(\d+)"\/>/g)].map((m) => parseInt(m[1], 10));
    expect(tcW.length).toBe(5);
    expect(tcW[3]).toBe(gridW[0] + gridW[1]); // AB2 spans cols 1+2
    // Table width converted from pct to explicit dxa (Word drag semantics).
    expect(xml).toContain('<w:tblW w:type="dxa"');
  });

  it("drag-moving a table stores page-anchored coordinates", async () => {
    const { moveTableTo } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    const tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table" || !tbl.src) throw new Error("not a table");

    expect(moveTableTo(doc, tbl.src, 120, 240)).toBe(true);

    const moved = doc.sections[0].blocks[0];
    if (moved.type !== "table") throw new Error("not a table");
    expect(moved.props.floating?.hAnchor).toBe("page");
    expect(moved.props.floating?.vAnchor).toBe("page");
    expect(moved.props.floating?.x).toBeCloseTo(120, 5);
    expect(moved.props.floating?.y).toBeCloseTo(240, 5);
    const xml = serializeXml(tbl.src);
    expect(xml).toContain('w:tblpX="1800"');
    expect(xml).toContain('w:tblpY="3600"');
  });

  it("keeps a table that started a later page on that page when drag-moving it", async () => {
    const { moveTableTo } = await import("../src/edit/tables.js");
    const doc = loadDoc(p("before") +
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
       <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
    const tbl = doc.sections[0].blocks[1];
    if (tbl.type !== "table" || !tbl.src) throw new Error("not a table");

    expect(moveTableTo(doc, tbl.src, 120, 96, true)).toBe(true);

    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toContain('<w:t xml:space="preserve">before</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p><w:tbl>');
    expect(xml).toContain('w:tblpY="1440"');
  });

  it("adds and removes page anchors when a moved table crosses pages", async () => {
    const { moveTableTo } = await import("../src/edit/tables.js");
    const doc = loadDoc(p("before") +
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
       <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
    const tbl = doc.sections[0].blocks[1];
    if (tbl.type !== "table" || !tbl.src) throw new Error("not a table");
    const pageBreaks = () => (serializeXml(doc.editableRoots()[0]).match(/<w:br w:type="page"\/>/g) ?? []).length;

    expect(moveTableTo(doc, tbl.src, 120, 96, false, 1)).toBe(true);
    expect(pageBreaks()).toBe(1);
    expect(moveTableTo(doc, tbl.src, 120, 96, false, 1)).toBe(true);
    expect(pageBreaks()).toBe(2);
    expect(moveTableTo(doc, tbl.src, 120, 96, false, -1)).toBe(true);
    expect(pageBreaks()).toBe(1);
  });

  it("adds Word's editable paragraph after terminal non-text blocks once", async () => {
    const { ensureParagraphAfterTerminalBlock } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
       <w:tr><w:tc><w:p><w:r><w:t>last cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    expect(doc.sections[0].blocks.map((block) => block.type)).toEqual(["table"]);
    expect(ensureParagraphAfterTerminalBlock(doc)).toBe(true);
    expect(doc.sections[0].blocks.map((block) => block.type)).toEqual(["table", "paragraph"]);
    expect(ensureParagraphAfterTerminalBlock(doc)).toBe(false);

    for (const terminal of [
      `<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`,
    ]) {
      const objectDoc = loadDoc(terminal);
      expect(ensureParagraphAfterTerminalBlock(objectDoc)).toBe(true);
      expect(objectDoc.sections[0].blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
      expect(ensureParagraphAfterTerminalBlock(objectDoc)).toBe(false);
    }
  });

  it("inserts and deletes rows", async () => {
    const { applyTableOp } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    expect(applyTableOp(doc, caretIn(doc, "A1"), "rowBelow")).toBe(true);
    let tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error("not a table");
    expect(tbl.rows.length).toBe(3);
    expect(applyTableOp(doc, caretIn(doc, "A2"), "deleteRow")).toBe(true);
    tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error("not a table");
    expect(tbl.rows.length).toBe(2);
  });

  it("inserts and deletes columns incl. grid", async () => {
    const { applyTableOp } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    expect(applyTableOp(doc, caretIn(doc, "B1"), "colRight")).toBe(true);
    let tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error("not a table");
    expect(tbl.rows[0].cells.length).toBe(3);
    expect(tbl.grid.length).toBe(3);
    expect(applyTableOp(doc, caretIn(doc, "A1"), "deleteCol")).toBe(true);
    tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error("not a table");
    expect(tbl.rows[0].cells.length).toBe(2);
    expect(tbl.grid.length).toBe(2);
  });

  it("deleting the last row removes the table", async () => {
    const { applyTableOp } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    applyTableOp(doc, caretIn(doc, "A2"), "deleteRow");
    applyTableOp(doc, caretIn(doc, "A1"), "deleteRow");
    expect(doc.sections[0].blocks.every((b) => b.type !== "table")).toBe(true);
  });

  it("Tab advances to the next cell, wrapping across rows", async () => {
    const { advanceCell } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    // A1 -> B1 (same row)
    expect(advanceCell(doc, caretIn(doc, "A1"), 1)?.t.text).toBe("B1");
    // B1 -> A2 (wrap to next row's first cell)
    expect(advanceCell(doc, caretIn(doc, "B1"), 1)?.t.text).toBe("A2");
  });

  it("Shift+Tab retreats to the previous cell; no-op in the first cell", async () => {
    const { advanceCell } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    expect(advanceCell(doc, caretIn(doc, "B1"), -1)?.t.text).toBe("A1");
    // A2 -> B1 (wrap back to previous row's last cell)
    expect(advanceCell(doc, caretIn(doc, "A2"), -1)?.t.text).toBe("B1");
    // First cell of the table has nowhere earlier to go.
    expect(advanceCell(doc, caretIn(doc, "A1"), -1)).toBeNull();
  });

  it("Tab in the last cell appends a new row and lands in its first cell", async () => {
    const { advanceCell } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    const dest = advanceCell(doc, caretIn(doc, "B2"), 1);
    const tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error("not a table");
    expect(tbl.rows.length).toBe(3); // a fresh row was appended
    expect(dest?.t.text).toBe(""); // caret in the new empty first cell
    // The new cell's w:t is real and typeable.
    dest!.t.text = "C1";
    const reloaded = DocxDocument.load(doc.save());
    const rt = reloaded.sections[0].blocks[0];
    if (rt.type !== "table") throw new Error("not a table");
    expect(rt.rows.length).toBe(3);
  });

  it("returns null outside any table", async () => {
    const { advanceCell } = await import("../src/edit/tables.js");
    const doc = tableDoc();
    expect(advanceCell(doc, caretIn(doc, "after"), 1)).toBeNull();
  });
});

describe("paragraph merge", () => {
  it("merges a paragraph into the previous one, keeping prev pPr", async () => {
    const { mergeParagraphBackward, paragraphOf } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>first</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>second</w:t></w:r></w:p>`,
    );
    const para2 = doc.sections[0].blocks[1] as Paragraph;
    const t2 = ((para2.children[0] as Run).content[0] as TextContent).srcT!;
    const pEl = paragraphOf(doc, t2 as never)!;
    expect(mergeParagraphBackward(doc, pEl)).toBe(true);
    expect(doc.sections[0].blocks.length).toBe(1);
    const merged = doc.sections[0].blocks[0] as Paragraph;
    expect(textOf(merged)).toBe("firstsecond");
    expect(merged.props.alignment).toBe("center"); // prev pPr wins
    const reloaded = DocxDocument.load(doc.save());
    expect(textOf(reloaded.sections[0].blocks[0] as Paragraph)).toBe("firstsecond");
  });

  it("refuses to merge across a table", async () => {
    const { mergeParagraphBackward, paragraphOf } = await import("../src/edit/blocks.js");
    const doc = loadDoc(
      p("before") +
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        p("after"),
    );
    const para = doc.sections[0].blocks[2] as Paragraph;
    const t = ((para.children[0] as Run).content[0] as TextContent).srcT!;
    const pEl = paragraphOf(doc, t as never)!;
    expect(mergeParagraphBackward(doc, pEl)).toBe(false);
  });
});

describe("image insertion", () => {
  it("adds media, relationship, content type, and inline drawing that round-trips", async () => {
    const { insertImageAt } = await import("../src/edit/blocks.js");
    const doc = loadDoc(p("caption here"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
    const relId = doc.addImageResource(png, "png");
    expect(relId).toMatch(/^rId\d+$/);
    expect(insertImageAt(doc, t as never, relId, 200, 100)).not.toBeNull();

    // model sees the image
    const para = doc.sections[0].blocks[0] as Paragraph;
    const imgs = para.children.flatMap((c) =>
      c.type === "run" ? c.content.filter((rc) => rc.kind === "image") : [],
    );
    expect(imgs.length).toBe(1);
    if (imgs[0].kind !== "image") throw new Error("not image");
    expect(imgs[0].width).toBeCloseTo(200, 0);
    expect(doc.media(imgs[0].part)).toBeDefined();

    // full zip round trip
    const reloaded = DocxDocument.load(doc.save());
    const rpara = reloaded.sections[0].blocks[0] as Paragraph;
    const rimgs = rpara.children.flatMap((c) =>
      c.type === "run" ? c.content.filter((rc) => rc.kind === "image") : [],
    );
    expect(rimgs.length).toBe(1);
    if (rimgs[0].kind !== "image") throw new Error("not image");
    expect(reloaded.media(rimgs[0].part)?.length).toBe(png.length);
    expect(reloaded.pkg.text("[Content_Types].xml")).toContain('Extension="png"');
  });

  it("stores SVG icons as native vector image parts", async () => {
    const { insertImageAt } = await import("../src/edit/blocks.js");
    const doc = loadDoc(p("icon"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><path d="M2 2h28v28H2z"/></svg>');
    const relId = doc.addImageResource(svg, "svg");
    expect(insertImageAt(doc, t, relId, 32, 32)).not.toBeNull();

    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.pkg.text("[Content_Types].xml")).toContain('Extension="svg" ContentType="image/svg+xml"');
    expect(reloaded.pkg.binary("word/media/image1.svg")).toEqual(svg);
  });
});

describe("chart insertion", () => {
  const DATA: ChartData = {
    type: "column",
    title: "Quarterly sales",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "Revenue", values: [12, 19, 15, 24] },
      { name: "Costs", values: [8, 11, 10, 14] },
    ],
  };

  it("packages native ChartML with an editable embedded workbook and round-trips cached data", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    expect(insertChartAt(doc, t, DATA)).not.toBeNull();

    const saved = DocxDocument.load(doc.save());
    expect(saved.pkg.names()).toContain("word/charts/chart1.xml");
    expect(saved.pkg.names()).toContain("word/charts/_rels/chart1.xml.rels");
    expect(saved.pkg.names()).toContain("word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    expect(saved.pkg.text("word/document.xml")).toContain("<c:chart");
    expect(saved.pkg.text("word/charts/chart1.xml")).toContain("Quarterly sales");
    expect(saved.pkg.text("word/charts/chart1.xml")).toContain("Data!$B$2:$B$5");
    expect(saved.pkg.text("word/charts/_rels/chart1.xml.rels")).toContain("relationships/package");
    expect(saved.pkg.text("[Content_Types].xml")).toContain("drawingml.chart+xml");

    const workbookBytes = saved.pkg.binary("word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    if (!workbookBytes) throw new Error("embedded workbook missing");
    const workbook = Package.from(workbookBytes);
    expect(workbook.names()).toContain("xl/worksheets/sheet1.xml");
    expect(workbook.text("xl/worksheets/sheet1.xml")).toContain("Revenue");
    expect(workbook.text("xl/worksheets/sheet1.xml")).toContain("<v>24</v>");

    const para = saved.sections[0].blocks[0] as Paragraph;
    const drawing = para.children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing") throw new Error("chart drawing missing");
    expect(drawing.chart).toEqual(DATA);
    expect(drawing.width).toBeCloseTo(480, 0);
    expect(drawing.height).toBeCloseTo(288, 0);
    expect(drawing.srcDrawing).toBeTruthy();
  });

  it("updates a selected chart's ChartML and workbook with undo and redo", async () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    insertChartAt(doc, t, DATA);
    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const drawing = para.children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing" || !drawing.srcDrawing) throw new Error("chart drawing missing");
    const source = drawing.srcDrawing;
    const updated: ChartData = {
      type: "line",
      title: "Updated trend",
      categories: ["Jan", "Feb", "Mar"],
      series: [{ name: "Total", values: [3, 7, 5] }],
    };
    const { EditHistory } = await import("../src/edit/history.js");
    const history = new EditHistory(reloaded);
    history.checkpoint();
    expect(setChartData(reloaded, source, updated)).toBe(true);
    expect(history.undo()).toBe(true);
    expect(history.lastTextChanges).toBeNull();
    expect(reloaded.pkg.text("word/charts/chart1.xml")).toContain("<c:barChart>");
    expect(history.redo()).toBe(true);
    expect(history.lastTextChanges).toBeNull();
    expect(reloaded.pkg.text("word/charts/chart1.xml")).toContain("<c:lineChart>");

    const saved = DocxDocument.load(reloaded.save());
    expect(saved.pkg.text("word/charts/chart1.xml")).toContain("<c:lineChart>");
    expect(saved.pkg.text("word/charts/chart1.xml")).toContain("Updated trend");
    const workbookBytes = saved.pkg.binary("word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    if (!workbookBytes) throw new Error("embedded workbook missing");
    expect(Package.from(workbookBytes).text("xl/worksheets/sheet1.xml")).toContain("<v>7</v>");
    const after = (saved.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!after || after.kind !== "drawing") throw new Error("updated chart missing");
    expect(after.srcDrawing).not.toBe(source);
    expect(after.chart).toEqual(updated);
  });

  it("emits the native plot element for every supported chart type", () => {
    expect(buildChartXml({ ...DATA, type: "column" })).toContain('<c:barDir val="col"/>');
    expect(buildChartXml({ ...DATA, type: "bar" })).toContain('<c:barDir val="bar"/>');
    expect(buildChartXml({ ...DATA, type: "line" })).toContain("<c:lineChart>");
    expect(buildChartXml({ ...DATA, type: "pie" })).toContain("<c:pieChart>");
  });
});

describe("SmartArt insertion", () => {
  const DATA: SmartArtData = { layout: "process", items: ["Discover", "Design", "Deliver"] };

  it("packages native diagram data/layout/drawing parts and round-trips editable nodes", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    expect(insertSmartArtAt(doc, t, DATA)).not.toBeNull();
    const saved = DocxDocument.load(doc.save());
    expect(saved.pkg.names()).toContain("word/diagrams/data1.xml");
    expect(saved.pkg.names()).toContain("word/diagrams/layout1.xml");
    expect(saved.pkg.names()).toContain("word/diagrams/quickStyle1.xml");
    expect(saved.pkg.names()).toContain("word/diagrams/colors1.xml");
    expect(saved.pkg.names()).toContain("word/diagrams/drawing1.xml");
    expect(saved.pkg.text("word/document.xml")).toContain("<dgm:relIds");
    expect(saved.pkg.text("word/diagrams/data1.xml")).toContain("Discover");
    expect(saved.pkg.text("word/diagrams/data1.xml")).toContain("dataModelExt");
    const drawingRelId = saved.pkg.text("word/diagrams/data1.xml").match(/dataModelExt[^>]+relId="([^"]+)"/)?.[1];
    expect(drawingRelId).toBeTruthy();
    expect(saved.pkg.text("word/_rels/document.xml.rels")).toContain(`Id="${drawingRelId}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"`);
    expect(saved.pkg.text("word/diagrams/layout1.xml")).toContain("urn:wordinweb:smartart:process");
    expect(saved.pkg.text("word/diagrams/layout1.xml")).toContain('<dgm:forEach axis="ch" ptType="node">');
    expect(saved.pkg.text("word/diagrams/layout1.xml")).toContain(
      '<dgm:presOf axis="desOrSelf" ptType="node"/>',
    );
    expect(saved.pkg.text("word/document.xml")).toContain('r:qs="rId3"');
    expect(saved.pkg.text("word/document.xml")).toContain('r:cs="rId4"');
    expect(saved.pkg.text("[Content_Types].xml")).toContain("drawingml.diagramData+xml");
    expect(saved.pkg.text("[Content_Types].xml")).toContain("drawingml.diagramStyle+xml");
    expect(saved.pkg.text("[Content_Types].xml")).toContain("drawingml.diagramColors+xml");
    expect(saved.pkg.text("[Content_Types].xml")).toContain("diagramDrawing+xml");

    const para = saved.sections[0].blocks[0] as Paragraph;
    const drawing = para.children.flatMap((item) => item.type === "run" ? item.content : []).find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing") throw new Error("SmartArt drawing missing");
    expect(drawing.smartArt).toEqual(DATA);
    expect(drawing.paths).toHaveLength(3);
    expect(drawing.texts).toHaveLength(3);
    expect(drawing.srcDrawing).toBeTruthy();
  });

  it("normalizes cached connector extents and keeps model ids tied to diagram data", () => {
    for (const layout of ["cycle", "hierarchy"] as const) {
      const xml = buildSmartArtDrawingXml({ layout, items: ["One", "Two", "Three"] });
      const extents = [...xml.matchAll(/<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/g)];
      expect(extents.length).toBeGreaterThan(0);
      for (const [, cx, cy] of extents) {
        expect(Number(cx)).toBeGreaterThan(0);
        expect(Number(cy)).toBeGreaterThan(0);
      }
      expect(xml).toMatch(/<a:xfrm[^>]+flip[HV]="1"/);
      for (const id of ["1", "2", "3"]) expect(xml).toContain(`modelId="${id}"`);
      const dataXml = buildSmartArtDataXml({ layout, items: ["One", "Two", "Three"] }, "rIdDrawing");
      for (const [, id] of xml.matchAll(/<dsp:sp modelId="([^"]+)"/g)) {
        expect(id).toMatch(/^\d+$/);
        expect(dataXml).toContain(`modelId="${id}"`);
      }
    }
  });

  it("updates selected SmartArt layout and text with undo and redo", async () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    insertSmartArtAt(doc, t, DATA);
    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const drawing = para.children.flatMap((item) => item.type === "run" ? item.content : []).find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing" || !drawing.srcDrawing) throw new Error("SmartArt drawing missing");
    const updated: SmartArtData = { layout: "hierarchy", items: ["Lead", "Plan", "Build", "Test"] };
    const { EditHistory } = await import("../src/edit/history.js");
    const history = new EditHistory(reloaded);
    history.checkpoint();
    expect(setSmartArtData(reloaded, drawing.srcDrawing, updated)).toBe(true);
    expect(reloaded.pkg.text("word/diagrams/data1.xml")).toContain("Lead");
    expect(reloaded.pkg.text("word/diagrams/layout1.xml")).toContain("smartart:hierarchy");
    expect(history.undo()).toBe(true);
    expect(reloaded.pkg.text("word/diagrams/data1.xml")).toContain("Discover");
    expect(history.redo()).toBe(true);
    expect(reloaded.pkg.text("word/diagrams/data1.xml")).toContain("Lead");
    const saved = DocxDocument.load(reloaded.save());
    const after = (saved.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!after || after.kind !== "drawing") throw new Error("updated SmartArt missing");
    expect(after.smartArt).toEqual(updated);
  });

  it("updates cached node fills without changing connector or text colors and round-trips", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    insertSmartArtAt(doc, t, { layout: "hierarchy", items: ["Lead", "Plan", "Build"] });
    const loaded = DocxDocument.load(doc.save());
    const drawing = (loaded.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing" || !drawing.srcDrawing) throw new Error("SmartArt drawing missing");
    const before = loaded.pkg.text("word/diagrams/drawing1.xml");
    const connectorColors = before.match(/<a:srgbClr val="7F8C8D"/g)?.length ?? 0;
    const textColors = before.match(/<a:srgbClr val="FFFFFF"/g)?.length ?? 0;

    expect(smartArtFillColor(loaded, drawing.srcDrawing, 1)).toBe("#ED7D31");
    expect(setSmartArtFill(loaded, drawing.srcDrawing, "#112233", 1)).toBe(true);
    expect(smartArtFillColor(loaded, drawing.srcDrawing, 0)).toBe("#4472C4");
    expect(smartArtFillColor(loaded, drawing.srcDrawing, 1)).toBe("#112233");
    expect(loaded.pkg.text("word/diagrams/drawing1.xml").match(/<a:srgbClr val="112233"/g)?.length).toBe(1);

    expect(setSmartArtFill(loaded, drawing.srcDrawing, "#AA22CC")).toBe(true);
    expect(smartArtFillColor(loaded, drawing.srcDrawing)).toBe("#AA22CC");
    const changed = loaded.pkg.text("word/diagrams/drawing1.xml");
    expect(changed.match(/<a:srgbClr val="AA22CC"/g)?.length).toBe(3);
    expect(changed.match(/<a:srgbClr val="7F8C8D"/g)?.length ?? 0).toBe(connectorColors);
    expect(changed.match(/<a:srgbClr val="FFFFFF"/g)?.length ?? 0).toBe(textColors);

    const saved = DocxDocument.load(loaded.save());
    const savedDrawing = (saved.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!savedDrawing || savedDrawing.kind !== "drawing" || !savedDrawing.srcDrawing) throw new Error("saved SmartArt drawing missing");
    expect(smartArtFillColor(saved, savedDrawing.srcDrawing)).toBe("#AA22CC");
    expect(saved.pkg.text("word/diagrams/drawing1.xml")).toContain('<a:srgbClr val="AA22CC"/>');
  });

  it("updates one node's editable and cached text and round-trips", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    insertSmartArtAt(doc, t, DATA);
    const loaded = DocxDocument.load(doc.save());
    const drawing = (loaded.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing" || !drawing.srcDrawing) throw new Error("SmartArt drawing missing");

    expect(setSmartArtNodeText(loaded, drawing.srcDrawing, 1, "Prototype & review")).toBe(true);
    expect(loaded.pkg.text("word/diagrams/data1.xml")).toContain("Prototype &amp; review");
    expect(loaded.pkg.text("word/diagrams/drawing1.xml")).toContain("Prototype &amp; review");
    expect(loaded.pkg.text("word/diagrams/data1.xml")).toContain("Discover");
    expect(loaded.pkg.text("word/diagrams/data1.xml")).toContain("Deliver");

    const saved = DocxDocument.load(loaded.save());
    const savedDrawing = (saved.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!savedDrawing || savedDrawing.kind !== "drawing") throw new Error("saved SmartArt missing");
    expect(savedDrawing.smartArt?.items).toEqual(["Discover", "Prototype & review", "Deliver"]);
  });

  it("formats one node's text in editable and cached SmartArt", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    insertSmartArtAt(doc, t, DATA);
    const loaded = DocxDocument.load(doc.save());
    const drawing = (loaded.sections[0].blocks[0] as Paragraph).children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "drawing");
    if (!drawing || drawing.kind !== "drawing" || !drawing.srcDrawing) throw new Error("SmartArt drawing missing");

    expect(setSmartArtTextFormat(loaded, drawing.srcDrawing, {
      fontFamily: "Arial",
      fontSizePt: 18,
      color: "#112233",
      bold: false,
      italic: true,
      alignment: "right",
    }, 1)).toBe(true);
    expect(smartArtTextFormat(loaded, drawing.srcDrawing, 1)).toEqual({
      fontFamily: "Arial",
      fontSizePt: 18,
      color: "#112233",
      bold: false,
      italic: true,
      alignment: "right",
    });
    const cached = loaded.pkg.text("word/diagrams/drawing1.xml");
    const editable = loaded.pkg.text("word/diagrams/data1.xml");
    expect(cached).toContain('sz="1800" b="0" i="1"');
    expect(cached).toContain('<a:latin typeface="Arial"/>');
    expect(cached).toContain('<a:srgbClr val="112233"/>');
    expect(editable).toContain('<a:latin typeface="Arial"/>');
  });

  it("builds cached geometry for every supported layout family", () => {
    for (const layout of ["process", "cycle", "hierarchy", "list"] as const) {
      const xml = buildSmartArtDrawingXml({ ...DATA, layout });
      expect(xml).toContain("<dsp:drawing");
      expect(xml).toContain("Discover");
      expect(xml).toContain("<a:prstGeom");
    }
  });
});

describe("advanced object insertion", () => {
  const POSTER = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

  const firstImage = (doc: DocxDocument) => {
    const para = doc.sections[0].blocks[0] as Paragraph;
    const image = para.children
      .flatMap((item) => item.type === "run" ? item.content : [])
      .find((item) => item.kind === "image");
    if (!image || image.kind !== "image") throw new Error("inserted object poster missing");
    return image;
  };

  it("packages a native 3D model with a selectable raster fallback", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    const glb = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);
    expect(insertModel3DAt(doc, t, { data: glb, poster: POSTER })).toBe(true);

    const saved = DocxDocument.load(doc.save());
    expect(saved.pkg.binary("word/media/model3d1.glb")).toEqual(glb);
    expect(saved.pkg.text("word/document.xml")).toContain("<am3d:model3d");
    expect(saved.pkg.text("word/document.xml")).toContain("<mc:AlternateContent");
    expect(saved.pkg.text("word/_rels/document.xml.rels")).toContain("relationships/model3d");
    expect(saved.pkg.text("[Content_Types].xml")).toContain('Extension="glb" ContentType="model/gltf-binary"');
    const image = firstImage(saved);
    expect(image.model3D).toEqual({ part: "word/media/model3d1.glb", posterPart: "word/media/image1.png" });
    expect(image.srcDrawing).toBeTruthy();
  });

  it("stores native online-video metadata without executing embedded HTML", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    expect(insertWebVideoAt(doc, t, { url: "https://youtu.be/dQw4w9WgXcQ", poster: POSTER })).toBe(true);

    const saved = DocxDocument.load(doc.save());
    const xml = saved.pkg.text("word/document.xml");
    expect(xml).toContain("<wp15:webVideoPr");
    expect(xml).toContain("https://www.youtube.com/embed/dQw4w9WgXcQ");
    const image = firstImage(saved);
    expect(image.webVideo?.url).toBe("https://www.youtube.com/embed/dQw4w9WgXcQ");
    expect(image.webVideo?.width).toBe(640);
    expect(image.srcDrawing).toBeTruthy();
  });

  it("round-trips arbitrary files through a native OLE Package and resizes its preview", () => {
    const payload = new TextEncoder().encode("embedded report data");
    const ole = buildOlePackage(payload, "report.txt");
    expect(Array.from(ole.slice(0, 8))).toEqual([208, 207, 17, 224, 161, 177, 26, 225]);
    const cfb = CFB.read(ole, { type: "array" });
    expect(cfb.FileIndex[0].clsid?.toUpperCase()).toBe("0C00030000000000C000000000000046");
    expect(Array.from(CFB.find(cfb, "\u0001Ole")?.content ?? [])).toEqual([
      1, 0, 0, 2,
      0, 0, 0, 0,
      0, 0, 0, 0,
      0, 0, 0, 0,
    ]);
    expect(extractOlePackage(ole)).toEqual({ filename: "report.txt", data: payload });

    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    expect(insertEmbeddedObjectAt(doc, t, { data: payload, filename: "report.txt", poster: POSTER })).toBe(true);
    let saved = DocxDocument.load(doc.save());
    const documentXml = saved.pkg.text("word/document.xml");
    expect(documentXml).toContain("<o:OLEObject");
    const oleObject = documentXml.match(/<o:OLEObject[^>]+\/>/)?.[0] ?? "";
    expect(oleObject).toMatch(/\bObjectID="_\d+"/);
    expect(oleObject).toContain('ProgID="Package"');
    const shapeId = oleObject.match(/\bShapeID="(_x0000_i(\d+))"/)?.[1];
    const shapeNumber = Number(oleObject.match(/\bShapeID="_x0000_i(\d+)"/)?.[1]);
    expect(shapeId).toBeTruthy();
    expect(shapeNumber).toBeGreaterThanOrEqual(1025);
    const shapeTypeId = documentXml.match(/<v:shapetype[^>]+\bid="(_x0000_t(\d+))"/)?.[1];
    const shapeTypeNumber = Number(documentXml.match(/<v:shapetype[^>]+\bid="_x0000_t(\d+)"/)?.[1]);
    expect(shapeTypeId).toBeTruthy();
    expect(shapeTypeNumber).toBeGreaterThanOrEqual(1025);
    expect(documentXml).toContain(`type="#${shapeTypeId}"`);
    expect(oleObject).toMatch(/\br:id="rId\d+"/);
    expect(documentXml.match(/<v:formulas>/g)).toHaveLength(1);
    expect(documentXml.match(/<v:f eqn=/g)).toHaveLength(12);
    expect(documentXml).toContain('<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>');
    expect(documentXml).toContain('<o:lock v:ext="edit" aspectratio="t"/>');
    expect(saved.pkg.text("word/_rels/document.xml.rels")).toContain("relationships/oleObject");
    expect(saved.pkg.text("[Content_Types].xml")).toContain('ContentType="application/vnd.openxmlformats-officedocument.oleObject"');
    const embedded = saved.pkg.binary("word/embeddings/oleObject1.bin");
    if (!embedded) throw new Error("OLE object part missing");
    expect(extractOlePackage(embedded)).toEqual({ filename: "report.txt", data: payload });

    const image = firstImage(saved);
    expect(image.embeddedObject).toEqual({
      part: "word/embeddings/oleObject1.bin",
      filename: "report.txt",
      progId: "Package",
    });
    expect(image.srcDrawing).toBeTruthy();
    expect(resizeDrawing(saved, image.srcDrawing!, 400, 240)).toBe(true);
    saved = DocxDocument.load(saved.save());
    expect(saved.pkg.text("word/document.xml")).toContain("width:300pt;height:180pt");
    expect(firstImage(saved).width).toBeCloseTo(400, 0);
    expect(firstImage(saved).height).toBeCloseTo(240, 0);
  });

  it("embeds DOCX bytes as an activatable Word.Document.12 package", () => {
    const embeddedDocx = loadDoc(p("Embedded source")).save();
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    expect(insertEmbeddedObjectAt(doc, t, {
      data: embeddedDocx,
      filename: "source.docx",
      poster: POSTER,
    })).toBe(true);

    const saved = DocxDocument.load(doc.save());
    const xml = saved.pkg.text("word/document.xml");
    expect(xml).toContain('ProgID="Word.Document.12"');
    expect(xml).toContain("<o:FieldCodes>\\s</o:FieldCodes>");
    expect(saved.pkg.text("word/_rels/document.xml.rels")).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Word_Document.docx"',
    );
    expect(saved.pkg.text("[Content_Types].xml")).toContain(
      'Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"',
    );
    expect(saved.pkg.binary("word/embeddings/Microsoft_Word_Document.docx")).toEqual(embeddedDocx);
    expect(firstImage(saved).embeddedObject).toEqual({
      part: "word/embeddings/Microsoft_Word_Document.docx",
      filename: "source.docx",
      progId: "Word.Document.12",
    });
  });

  it("repairs legacy WordInWeb SmartArt and DOCX Package objects on load", () => {
    const embeddedDocx = loadDoc(p("Legacy embedded source")).save();
    const source = loadDoc(p("Anchor"));
    const t = (firstRun(source).run.content[0] as TextContent).srcT!;
    insertSmartArtAt(source, t, { layout: "cycle", items: ["Discover", "Design", "Deliver"] });
    insertEmbeddedObjectAt(source, t, { data: embeddedDocx, filename: "legacy.docx", poster: POSTER });
    const files = unzipSync(source.save());

    files["word/diagrams/data1.xml"] = strToU8(
      strFromU8(files["word/diagrams/data1.xml"])
        .replace('modelId="4" srcId=', 'modelId="c1" srcId=')
        .replace('modelId="5" srcId=', 'modelId="c2" srcId=')
        .replace('modelId="6" srcId=', 'modelId="c3" srcId='),
    );
    files["word/diagrams/drawing1.xml"] = strToU8(
      strFromU8(files["word/diagrams/drawing1.xml"]).replace(/<a:ext cx="2841862" cy="1"\/>/, '<a:ext cx="-2841862" cy="0"/>'),
    );
    const documentXml = strFromU8(files["word/document.xml"])
      .replace(/_x0000_t\d+/g, "_x0000_t75_1002")
      .replace(/_x0000_i\d+/g, "_x0000_i1002")
      .replace('ProgID="Word.Document.12"', 'ProgID="Package"')
      .replace("<o:FieldCodes>\\s</o:FieldCodes>", "");
    files["word/document.xml"] = strToU8(documentXml);
    files["word/_rels/document.xml.rels"] = strToU8(
      strFromU8(files["word/_rels/document.xml.rels"]).replace(
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Word_Document.docx"',
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"',
      ),
    );
    files["word/embeddings/oleObject1.bin"] = buildOlePackage(embeddedDocx, "legacy.docx");
    delete files["word/embeddings/Microsoft_Word_Document.docx"];
    files["[Content_Types].xml"] = strToU8(
      strFromU8(files["[Content_Types].xml"])
        .replace(/<Default Extension="docx"[^>]+\/>/, "")
        .replace("</Types>", '<Override PartName="/word/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>'),
    );

    const repaired = DocxDocument.load(zipSync(files));
    const saved = DocxDocument.load(repaired.save());
    const repairedData = saved.pkg.text("word/diagrams/data1.xml");
    const repairedDrawing = saved.pkg.text("word/diagrams/drawing1.xml");
    expect(repairedData).not.toMatch(/modelId="c\d+"/);
    expect(repairedDrawing).not.toMatch(/<a:ext cx="-?\d+" cy="(?:0|-\d+)"/);
    expect(repairedDrawing).toContain("Discover");
    expect(saved.pkg.text("word/document.xml")).toContain('ProgID="Word.Document.12"');
    expect(saved.pkg.text("word/document.xml")).toContain("<o:FieldCodes>\\s</o:FieldCodes>");
    expect(saved.pkg.text("word/_rels/document.xml.rels")).toContain("relationships/package");
    expect(saved.pkg.binary("word/embeddings/Microsoft_Word_Document.docx")).toEqual(embeddedDocx);
    expect(saved.pkg.has("word/embeddings/oleObject1.bin")).toBe(false);
  });
});

describe("blank page insertion", () => {
  it("splits head and tail around a distinct editable blank-page paragraph", () => {
    const doc = loadDoc(
      `<w:p w:rsidR="1"><w:pPr><w:spacing w:after="120"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr>` +
      `<w:bookmarkStart w:id="0" w:name="Split"/><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Alpha</w:t><w:tab/></w:r><w:bookmarkEnd w:id="0"/></w:p>`,
    );
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    expect(insertBlankPageAt(doc, t, 2)).toBe(true);
    const saved = doc.save();
    const reloaded = DocxDocument.load(saved);
    const xml = reloaded.pkg.text("word/document.xml");
    expect((xml.match(/<w:br w:type="page"\/>/g) ?? [])).toHaveLength(2);
    expect((xml.match(/<w:sectPr>/g) ?? [])).toHaveLength(1);
    const paragraphs = [...xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
    expect(paragraphs).toHaveLength(3);
    expect(paragraphs[0]).toContain("Al");
    expect(paragraphs[0]).toContain("bookmarkStart");
    expect(paragraphs[0]).not.toContain("sectPr");
    expect(paragraphs[1]).toMatch(/<w:t xml:space="preserve"\/><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:br w:type="page"\/>/);
    expect(paragraphs[1]).not.toContain("sectPr");
    expect(paragraphs[2]).toContain("pha");
    expect(paragraphs[2]).toContain("<w:tab/>");
    expect(paragraphs[2]).toContain("bookmarkEnd");
    expect(paragraphs[2]).toContain("<w:sectPr>");
    const blocks = reloaded.sections.flatMap((section) => section.blocks) as Paragraph[];
    expect(blocks.map(textOf)).toEqual(["Al", "", "pha"]);
    const layout = layoutDocument(reloaded);
    expect(layout.pages).toHaveLength(3);
    expect(
      layout.pages[1].items.some((item) => item.kind === "text" && item.text === "" && item.src?.t),
    ).toBe(true);
  });

  it("keeps a blank page independent when inserted inside a bulleted list", () => {
    const doc = loadDoc(
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t>Alpha</w:t></w:r></w:p>`,
    );
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;

    expect(insertBlankPageAt(doc, t, 2)).toBe(true);

    const xml = DocxDocument.load(doc.save()).pkg.text("word/document.xml");
    const paragraphs = [...xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
    expect(paragraphs).toHaveLength(3);
    expect(paragraphs[0]).toContain("<w:numPr>");
    expect(paragraphs[1]).not.toContain("<w:numPr>");
    expect(paragraphs[2]).toContain("<w:numPr>");
  });
});

describe("page and section navigation", () => {
  it("returns an editable target after a trailing page break", () => {
    const doc = loadDoc(p("Plain"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    const destination = insertBreakAt(doc, t, t.text.length, "page");

    expect(destination).not.toBeNull();
    expect(destination?.offset).toBe(0);
    expect(destination?.t.text).toBe("");
    expect(doc.findParentOf(destination!.t)).toBeTruthy();
    expect(layoutDocument(doc).pages).toHaveLength(2);
  });

  it("reports the logical section at the caret when continuous sections share a page", () => {
    const continuous = `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const sectionParagraph = (text: string) =>
      `<w:p><w:pPr>${continuous}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const doc = loadDoc(sectionParagraph("Intro") + sectionParagraph("Article") + p("Ending") + continuous);
    const texts = doc.sections.map((section) => {
      const paragraph = section.blocks[0] as Paragraph;
      return (paragraph.children[0] as Run).content.find((content) => content.kind === "text") as TextContent;
    });

    expect(sectionContextAt(doc, texts[0].srcT!)).toEqual({ index: 1, count: 3 });
    expect(sectionContextAt(doc, texts[1].srcT!)).toEqual({ index: 2, count: 3 });
    expect(sectionContextAt(doc, texts[2].srcT!)).toEqual({ index: 3, count: 3 });
  });
});

describe("cover page insertion", () => {
  it("prepends editable title, subtitle, author, and a page break", () => {
    const doc = loadDoc(p("Existing content"));
    expect(insertCoverPage(doc, { title: "Project Atlas", subtitle: "Launch plan", author: "Ada Lovelace" })).toBe(true);
    const saved = doc.save();
    const xml = DocxDocument.load(saved).pkg.text("word/document.xml");
    expect(xml).toContain("Project Atlas");
    expect(xml).toContain("Launch plan");
    expect(xml).toContain("Ada Lovelace");
    expect(xml).toContain('w:pStyle w:val="Title"');
    expect(xml.indexOf("Project Atlas")).toBeLessThan(xml.indexOf("Existing content"));
    expect((xml.match(/<w:br w:type="page"\/>/g) ?? [])).toHaveLength(1);
  });
});

describe("shape insertion", () => {
  it("creates a native vertical line and preserves its axis when resized", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const drawing = insertShapeAt(doc, t, "verticalLine")!;

    expect(serializeXml(drawing)).toContain('<a:prstGeom prst="line"');
    expect(serializeXml(drawing)).toContain('<a:ext cx="0" cy="2286000"/>');
    expect(resizeDrawing(doc, drawing, 4, 500)).toBe(true);
    expect(serializeXml(drawing)).toContain('<wp:extent cx="38100" cy="4762500"/>');
    expect(serializeXml(drawing)).toContain('<a:ext cx="0" cy="4762500"/>');

    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const anchor = para.children.flatMap((child) => child.type === "run" ? child.content : [])
      .find((content) => content.kind === "anchor");
    if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "art") throw new Error("vertical line missing");
    expect(anchor.shape.lines).toHaveLength(1);
    expect(anchor.shape.lines[0].x1).toBe(anchor.shape.lines[0].x2);
    expect(anchor.shape.lines[0].y2 - anchor.shape.lines[0].y1).toBeCloseTo(500, 1);
  });

  it("creates native anchored DrawingML shapes that retain editable text", () => {
    const presets: ShapePreset[] = ["rectangle", "roundedRectangle", "ellipse", "diamond", "textBox"];
    for (const preset of presets) {
      const doc = loadDoc(p("Anchor"));
      const { run } = firstRun(doc);
      const t = (run.content[0] as TextContent).srcT!;
      expect(insertShapeAt(doc, t, preset, "Editable shape")).not.toBeNull();
      const xml = DocxDocument.load(doc.save()).pkg.text("word/document.xml");
      expect(xml).toContain("<wp:anchor");
      expect(xml).toContain("<wps:wsp");
      expect(xml).toContain("<w:txbxContent>");
      expect(xml).toContain("Editable shape");
      const expectedGeom = preset === "roundedRectangle" ? "roundRect" : preset === "textBox" || preset === "rectangle" ? "rect" : preset;
      expect(xml).toContain(`a:prstGeom prst="${expectedGeom}"`);

      const reloaded = DocxDocument.load(doc.save());
      const para = reloaded.sections[0].blocks[0] as Paragraph;
      const anchor = para.children.flatMap((child) => child.type === "run" ? child.content : []).find((content) => content.kind === "anchor");
      expect(anchor?.kind).toBe("anchor");
      if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "textbox") throw new Error("shape missing");
      expect(anchor.shape.blocks[0]?.type).toBe("paragraph");
      expect(anchor.shape.srcDrawing).toBeTruthy();
      expect(anchor.shape.textboxStory).toBe(true);
    }
  });

  it("keeps an empty text box as a real selectable DrawingML object", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const drawing = insertShapeAt(doc, t, "textBox");
    expect(drawing).not.toBeNull();
    const xml = serializeXml(drawing!);
    expect(xml).toContain("<a:noFill/>");
    expect(xml).toContain("<a:ln");
    expect(xml).toContain("<w:t");
  });

  it("adds an explicit outline when a shape only has Word's default line style", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const drawing = insertShapeAt(doc, t, "rectangle", "Styled")!;
    const find = (element: XmlElement, name: string): XmlElement | undefined => {
      if (localName(element.name) === name) return element;
      for (const child of element.children) {
        const match = find(child, name);
        if (match) return match;
      }
      return undefined;
    };
    const spPr = find(drawing, "spPr")!;
    spPr.children = spPr.children.filter((child) => localName(child.name) !== "ln");

    expect(drawingLineStyle(drawing)).toEqual({ color: "#000000", width: 0.75, dash: "solid" });
    expect(setDrawingLineStyle(doc, drawing, "#00FF00", 3, "dashed")).toBe(true);
    expect(serializeXml(drawing)).toContain('<a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/>');
    expect(serializeXml(drawing)).toContain('<a:prstDash val="dash"/>');
    expect(drawingLineStyle(drawing)).toEqual({ color: "#00FF00", width: 3, dash: "dashed" });
  });

  it("persists page alignment, rotation, and front/back order", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const first = insertShapeAt(doc, t, "rectangle", "First")!;
    const second = insertShapeAt(doc, t, "ellipse", "Second")!;

    expect(setFloatingPagePosition(doc, first, 120, 240)).toBe(true);
    expect(setDrawingRotation(doc, first, 90)).toBe(true);
    expect(drawingRotation(first)).toBe(90);
    expect(setDrawingOrder(doc, first, "front")).toBe(true);
    const frontXml = serializeXml(first);
    const secondXml = serializeXml(second);
    expect(frontXml).toContain('relativeFrom="page"');
    expect(frontXml).toContain("1143000");
    expect(frontXml).toContain("2286000");
    expect(frontXml).toContain('rot="5400000"');
    expect(Number(/relativeHeight="(\d+)"/.exec(frontXml)![1])).toBeGreaterThan(
      Number(/relativeHeight="(\d+)"/.exec(secondXml)![1]),
    );
    expect(serializeXml(doc.docRoot).indexOf("First")).toBeGreaterThan(serializeXml(doc.docRoot).indexOf("Second"));

    expect(setDrawingOrder(doc, first, "back")).toBe(true);
    expect(serializeXml(first)).toContain('relativeHeight="0"');
    expect(serializeXml(doc.docRoot).indexOf("First")).toBeLessThan(serializeXml(doc.docRoot).indexOf("Second"));
    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const shape = para.children
      .flatMap((child) => child.type === "run" ? child.content : [])
      .find((content) => content.kind === "anchor" && content.shape.type === "textbox" && content.shape.rotation === 90);
    if (!shape || shape.kind !== "anchor" || shape.shape.type !== "textbox") throw new Error("shape missing");
    expect(shape.shape.rotation).toBe(90);
  });
});

describe("WordArt insertion", () => {
  it("creates native editable DrawingML text warps that round-trip", () => {
    const presets: Array<[WordArtPreset, string]> = [
      ["plain", "textNoShape"],
      ["archUp", "textArchUp"],
      ["archDown", "textArchDown"],
      ["wave", "textWave1"],
      ["chevron", "textChevron"],
    ];
    for (const [preset, warp] of presets) {
      const doc = loadDoc(p("Anchor"));
      const { run } = firstRun(doc);
      const t = (run.content[0] as TextContent).srcT!;
      expect(insertWordArtAt(doc, t, "Editable WordArt", preset)).not.toBeNull();
      const saved = doc.save();
      const xml = DocxDocument.load(saved).pkg.text("word/document.xml");
      expect(xml).toContain('name="WordArt ');
      expect(xml).toContain(`<a:prstTxWarp prst="${warp}">`);
      expect(xml).toContain("Editable WordArt");
      expect(xml).toContain('<w:color w:val="2E74B5"');

      const reloaded = DocxDocument.load(saved);
      const para = reloaded.sections[0].blocks[0] as Paragraph;
      const anchor = para.children.flatMap((child) => child.type === "run" ? child.content : []).find((content) => content.kind === "anchor");
      if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "textbox") throw new Error("WordArt missing");
      expect(anchor.shape.warp).toBe(warp === "textNoShape" ? undefined : warp);
      expect(anchor.shape.blocks[0]?.type).toBe("paragraph");
    }
  });

  it("does not insert empty WordArt", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    expect(insertWordArtAt(doc, t, "", "plain")).toBeNull();
  });

  it("uses tight plain-text bounds and edits DrawingML text through save/reopen", () => {
    const doc = loadDoc(p("Anchor"));
    const t = (firstRun(doc).run.content[0] as TextContent).srcT!;
    const drawing = insertWordArtAt(doc, t, "Before", "plain");
    if (!drawing) throw new Error("WordArt missing");
    expect(isDrawingWordArt(drawing)).toBe(true);
    expect(drawingWordArtText(drawing)).toBe("Before");
    expect(serializeXml(drawing)).toContain('cy="381000"');
    expect(setDrawingWordArtText(doc, drawing, "After")).toBe(true);
    const saved = DocxDocument.load(doc.save());
    const xml = saved.pkg.text("word/document.xml");
    expect(xml).toContain("After");
    expect(xml).not.toContain(">Before<");
    expect(xml).toContain('cy="381000"');
  });
});

describe("ink insertion", () => {
  it("creates movable freehand DrawingML geometry that round-trips", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    expect(insertInkAt(doc, t, [
      { x: 120, y: 180 },
      { x: 140, y: 195 },
      { x: 165, y: 178 },
    ], "#C00000", 4, 0.45)).not.toBeNull();

    const saved = doc.save();
    const xml = DocxDocument.load(saved).pkg.text("word/document.xml");
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain("<a:custGeom>");
    expect(xml).toContain("<a:moveTo>");
    expect(xml).toContain("<a:lnTo>");
    expect(xml).toContain('a:srgbClr val="C00000"');
    expect(xml).toContain('<a:alpha val="45000"/>');

    const reloaded = DocxDocument.load(saved);
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const anchor = para.children.flatMap((child) => child.type === "run" ? child.content : []).find((content) => content.kind === "anchor");
    expect(anchor?.kind).toBe("anchor");
    if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "art") throw new Error("ink missing");
    expect(anchor.shape.ink).toBe(true);
    expect(anchor.shape.paths).toHaveLength(1);
    expect(anchor.shape.paths[0].d).toMatch(/^M .* L .* L /);
    expect(anchor.shape.paths[0].stroke?.opacity).toBe(0.45);
    expect(anchor.shape.srcDrawing).toBeTruthy();
  });

  it("rotates native ink geometry as one selectable object", () => {
    const doc = loadDoc(p("Anchor"));
    const { run } = firstRun(doc);
    const t = (run.content[0] as TextContent).srcT!;
    const drawing = insertInkAt(doc, t, [{ x: 20, y: 30 }, { x: 80, y: 60 }])!;
    expect(setDrawingRotation(doc, drawing, 90)).toBe(true);
    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const anchor = para.children.flatMap((child) => child.type === "run" ? child.content : []).find((content) => content.kind === "anchor");
    if (!anchor || anchor.kind !== "anchor" || anchor.shape.type !== "art") throw new Error("ink missing");
    expect(anchor.shape.rotation).toBe(90);
  });
});

describe("addComment", () => {
  it("creates comments.xml on demand, anchors the range, and round-trips through save", async () => {
    const doc = loadDoc(p("Hello brave new world"));
    const { run } = firstRun(doc);
    const ok = addComment(doc, [segFor(run, 6, 11)], "Check this", "Reviewer", "RV");
    expect(ok).toBe(true);
    expect(doc.comments.length).toBe(1);
    expect(doc.comments[0].author).toBe("Reviewer");
    // range anchors around the selected word
    const anchors = doc.commentAnchors();
    const ts = anchors.get(doc.comments[0].id) ?? [];
    expect(ts.map((t) => t.text).join("")).toBe("brave");
    // survives save/reload including the created part + rels + content types
    const doc2 = DocxDocument.load(doc.save());
    expect(doc2.comments.length).toBe(1);
    expect(doc2.comments[0].author).toBe("Reviewer");
    const ts2 = doc2.commentAnchors().get(doc2.comments[0].id) ?? [];
    expect(ts2.map((t) => t.text).join("")).toBe("brave");
  });
});

describe("setListType", () => {
  it("creates numbering.xml on demand, renders a bullet label, and round-trips", () => {
    const doc = loadDoc(p("alpha") + p("beta"));
    const { run } = firstRun(doc);
    const t = run.content.find((c): c is TextContent => c.kind === "text")!.srcT!;
    expect(setListType(doc, [t], "bullet")).toBe(true);
    const para = doc.sections[0].blocks[0] as Paragraph;
    expect(para.props.numbering?.numId).toBeGreaterThan(0);
    // survives save/reload including the created part + rels + content types
    const doc2 = DocxDocument.load(doc.save());
    const para2 = doc2.sections[0].blocks[0] as Paragraph;
    expect(para2.props.numbering?.numId).toBe(para.props.numbering?.numId);
    const lvl = doc2.numberingLevel(para2.props.numbering!.numId, 0);
    expect(lvl?.format).toBe("bullet");
    // toggle off
    expect(setListType(doc, [t], null)).toBe(true);
    expect((doc.sections[0].blocks[0] as Paragraph).props.numbering).toBeUndefined();
  });
});

describe("hyperlinks", () => {
  it("wraps a selection, reports and retargets, and unwraps", () => {
    const doc = loadDoc(p("Visit our website today"));
    const { run } = firstRun(doc);
    expect(setLink(doc, [segFor(run, 6, 17)], "https://a.example")).toBe(true);
    const para = doc.sections[0].blocks[0] as Paragraph;
    const link = para.children.find((c) => c.type === "hyperlink");
    expect(link && link.type === "hyperlink" ? link.href : null).toBe("https://a.example");
    expect(textOf(para)).toBe("Visit our website today");
    // linkAt through the covered w:t
    const linkT = (link && link.type === "hyperlink" ? link.runs[0].content[0] : null);
    const tEl = linkT && linkT.kind === "text" ? linkT.srcT! : null;
    expect(linkAt(doc, tEl!)).toBe("https://a.example");
    // retarget
    expect(setLink(doc, [{ run: (link as { runs: Run[] }).runs[0], t: tEl, start: 0, end: 3, props: {} }], "https://b.example")).toBe(true);
    expect(linkAt(doc, tEl!)).toBe("https://b.example");
    // round-trip
    const doc2 = DocxDocument.load(doc.save());
    const para2 = doc2.sections[0].blocks[0] as Paragraph;
    const link2 = para2.children.find((c) => c.type === "hyperlink");
    expect(link2 && link2.type === "hyperlink" ? link2.href : null).toBe("https://b.example");
    // unwrap
    expect(removeLink(doc, tEl!)).toBe(true);
    const para3 = doc.sections[0].blocks[0] as Paragraph;
    expect(para3.children.every((c) => c.type === "run")).toBe(true);
    expect(textOf(para3)).toBe("Visit our website today");
  });
});

describe("paragraph formatting", () => {
  it("indents and outdents in half-inch steps with a floor at zero", () => {
    const doc = loadDoc(p("target"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(adjustIndent(doc, [t], 1)).toBe(true);
    expect((doc.sections[0].blocks[0] as Paragraph).props.indentLeft).toBeCloseTo(48, 1); // 720tw = 48px
    expect(adjustIndent(doc, [t], -1)).toBe(true);
    expect((doc.sections[0].blocks[0] as Paragraph).props.indentLeft ?? 0).toBe(0);
    expect(adjustIndent(doc, [t], -1)).toBe(false); // floored
  });

  it("sets line spacing and space before/after", () => {
    const doc = loadDoc(p("target"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(setParagraphSpacing(doc, [t], { lineMultiple: 1.5, afterPt: 12 })).toBe(true);
    const props = (doc.sections[0].blocks[0] as Paragraph).props;
    expect(props.lineSpacing?.rule).toBe("auto");
    expect(props.lineSpacing?.value).toBeCloseTo(1.5, 2);
  });

  it("creates, reads, customizes, and removes a paragraph divider", () => {
    const doc = loadDoc(p("target"));
    const { run } = firstRun(doc);
    const t = run.content.find((content) => content.kind === "text")!.srcT!;

    expect(paragraphDividerAt(doc, t)).toBeNull();
    expect(setParagraphDivider(doc, [t], {
      style: "thinThickSmallGap",
      color: "#2E74B5",
      widthPt: 3,
      spacePt: 1,
    })).toBe(true);
    expect(paragraphDividerAt(doc, t)).toEqual({
      style: "thinThickSmallGap",
      color: "#2E74B5",
      widthPt: 3,
      spacePt: 1,
    });
    expect(serializeXml(doc.editableRoots()[0])).toContain(
      '<w:pBdr><w:bottom w:val="thinThickSmallGap" w:sz="24" w:space="1" w:color="2E74B5"/></w:pBdr>',
    );

    expect(setParagraphDivider(doc, [t], null)).toBe(true);
    expect(paragraphDividerAt(doc, t)).toBeNull();
    expect(serializeXml(doc.editableRoots()[0])).not.toContain("<w:pBdr>");
  });

  it("sets and round-trips an exact line height in points", () => {
    const doc = loadDoc(p("target"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(setParagraphSpacing(doc, [t], { exactLinePt: 24, beforePt: 0, afterPt: 0 })).toBe(true);

    const spacing = (doc.sections[0].blocks[0] as Paragraph).props.lineSpacing;
    expect(spacing?.rule).toBe("exact");
    expect(spacing?.value).toBeCloseTo(32, 1); // 24pt at 96 CSS pixels per inch

    const saved = doc.save();
    const xml = strFromU8(unzipSync(saved)["word/document.xml"]);
    expect(xml).toContain('w:line="480"');
    expect(xml).toContain('w:lineRule="exact"');
    expect(xml).toContain('w:before="0"');
    expect(xml).toContain('w:after="0"');
    const reloaded = DocxDocument.load(saved);
    expect((reloaded.sections[0].blocks[0] as Paragraph).props.lineSpacing).toMatchObject({ rule: "exact" });
  });

  it("applies, changes, removes, and round-trips a native drop cap", () => {
    const doc = loadDoc(p("Once upon a time"));
    const { run } = firstRun(doc);
    const t = run.content.find((content) => content.kind === "text")!.srcT!;
    expect(setDropCapAt(doc, t, "drop", 3)).toBe(true);
    expect(doc.sections[0].blocks).toHaveLength(2);
    const cap = doc.sections[0].blocks[0] as Paragraph;
    const body = doc.sections[0].blocks[1] as Paragraph;
    expect(cap.props.dropCap).toMatchObject({ mode: "drop", lines: 3 });
    expect(textOf(cap)).toBe("O");
    expect(textOf(body)).toBe("nce upon a time");

    const bodyT = (body.children[0] as Run).content.find((content) => content.kind === "text")!.srcT!;
    expect(setDropCapAt(doc, bodyT, "margin", 2)).toBe(true);
    expect((doc.sections[0].blocks[0] as Paragraph).props.dropCap).toMatchObject({ mode: "margin", lines: 2, pageAnchored: true });
    expect(DocxDocument.load(doc.save()).sections[0].blocks).toHaveLength(2);

    const changedBody = doc.sections[0].blocks[1] as Paragraph;
    const changedT = (changedBody.children[0] as Run).content.find((content) => content.kind === "text")!.srcT!;
    expect(setDropCapAt(doc, changedT, null)).toBe(true);
    expect(doc.sections[0].blocks).toHaveLength(1);
    expect(textOf(doc.sections[0].blocks[0] as Paragraph)).toBe("Once upon a time");
  });
});

describe("find & replace / case", () => {
  it("finds across split runs and replaces preserving surroundings", () => {
    const doc = loadDoc(p("The cat sat on the cat mat"));
    const { run } = firstRun(doc);
    // split "cat" across runs by bolding half of the first one
    applyRunFormat(doc, [segFor(run, 4, 6)], { bold: true });
    const hits = findAll(doc, "cat");
    expect(hits.length).toBe(2);
    expect(hits[0].ranges.length).toBeGreaterThan(1); // spans split runs
    const n = replaceAll(doc, "cat", "dog");
    expect(n).toBe(2);
    expect(textOf(doc.sections[0].blocks[0] as Paragraph)).toBe("The dog sat on the dog mat");
  });

  it("changes case over a selection", () => {
    const doc = loadDoc(p("make me shout"));
    const { run } = firstRun(doc);
    transformCase(doc, [segFor(run, 0, 13)], "upper");
    expect(textOf(doc.sections[0].blocks[0] as Paragraph)).toBe("MAKE ME SHOUT");
  });

  it("clear formatting strips direct run formatting", () => {
    const doc = loadDoc(p("styled text"));
    const { run } = firstRun(doc);
    applyRunFormat(doc, [segFor(run, 0, 11)], { bold: true, color: "#FF0000" });
    const { run: run2 } = firstRun(doc);
    applyRunFormat(doc, [segFor(run2, 0, 11)], { clear: true });
    const para = doc.sections[0].blocks[0] as Paragraph;
    const r = para.children[0];
    expect(r.type === "run" ? r.props.bold : true).toBeUndefined();
  });
});

describe("list levels", () => {
  it("steps ilvl with Tab/Shift-Tab semantics, clamped", () => {
    const doc = loadDoc(p("item"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    setListType(doc, [t], "bullet");
    expect(setListLevel(doc, [t], 1)).toBe(true);
    expect((doc.sections[0].blocks[0] as Paragraph).props.numbering?.ilvl).toBe(1);
    expect(setListLevel(doc, [t], -1)).toBe(true);
    expect(setListLevel(doc, [t], -1)).toBe(false); // clamped at 0
  });
});

describe("cell merge/split", () => {
  const TBL = `<w:tbl>
    <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>
      <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
    </w:tr>
    <w:tr>
      <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
    </w:tr>
  </w:tbl>`;
  const tOf = (doc: DocxDocument, text: string): XmlElement => {
    let found: XmlElement | undefined;
    const walk = (e: XmlElement) => {
      if (e.name.endsWith("t") && e.text === text) found = e;
      for (const c of e.children) walk(c);
    };
    for (const root of doc.editableRoots()) walk(root);
    if (!found) throw new Error("t not found: " + text);
    return found;
  };

  it("merges right (gridSpan + content) and splits back", () => {
    const doc = loadDoc(TBL + p("after"));
    expect(applyTableOp(doc, tOf(doc, "A1"), "mergeRight")).toBe(true);
    let tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error();
    expect(tbl.rows[0].cells.length).toBe(1);
    expect(tbl.rows[0].cells[0].props.gridSpan).toBe(2);
    expect(applyTableOp(doc, tOf(doc, "A1"), "splitCell")).toBe(true);
    tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error();
    expect(tbl.rows[0].cells.length).toBe(2);
  });

  it("merges down (vMerge) and splits back", () => {
    const doc = loadDoc(TBL + p("after"));
    expect(applyTableOp(doc, tOf(doc, "A1"), "mergeDown")).toBe(true);
    let tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error();
    expect(tbl.rows[0].cells[0].props.vMerge).toBe("restart");
    expect(tbl.rows[1].cells[0].props.vMerge).toBe("continue");
    expect(applyTableOp(doc, tOf(doc, "A1"), "splitCell")).toBe(true);
    tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error();
    expect(tbl.rows[0].cells[0].props.vMerge).toBeUndefined();
    expect(tbl.rows[1].cells[0].props.vMerge).toBeUndefined();
  });

  it("sets cell shading and vertical alignment", () => {
    const doc = loadDoc(TBL + p("after"));
    expect(cellShadingAt(doc, tOf(doc, "B2"))).toBeNull();
    expect(applyTableOp(doc, tOf(doc, "B2"), { kind: "cellShading", fill: "FFF2CC" })).toBe(true);
    expect(cellShadingAt(doc, tOf(doc, "B2"))).toBe("#FFF2CC");
    expect(cellShadingAt(doc, tOf(doc, "after"))).toBeUndefined();
    expect(applyTableOp(doc, tOf(doc, "B2"), { kind: "cellVAlign", v: "center" })).toBe(true);
    const tbl = doc.sections[0].blocks[0];
    if (tbl.type !== "table") throw new Error();
    expect(tbl.rows[1].cells[1].props.shading?.toUpperCase()).toContain("FFF2CC");
    expect(tbl.rows[1].cells[1].props.verticalAlign).toBe("center");
  });
});

describe("image editing", () => {
  const DRAWING = `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="914400" cy="914400"/>
      <wp:docPr id="1" name="Pic" descr="old alt"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:blipFill><a:blip r:embed="rIdIMG" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:srcRect l="25000" t="10000"/></pic:blipFill>
            <pic:spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
  const RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const DOCRELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdIMG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;
  const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  function loadWithImage() {
    return DocxDocument.load(
      makeDocxWithMedia(
        { "word/document.xml": wrapDocument(DRAWING), "word/_rels/document.xml.rels": DOCRELS, "_rels/.rels": RELS },
        { "word/media/image1.png": PNG },
      ),
    );
  }

  it("adds a trailing editable paragraph after a terminal image-only paragraph", async () => {
    const { ensureParagraphAfterTerminalBlock } = await import("../src/edit/blocks.js");
    const doc = loadWithImage();
    expect(ensureParagraphAfterTerminalBlock(doc)).toBe(true);
    expect(doc.sections[0].blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect(ensureParagraphAfterTerminalBlock(doc)).toBe(false);
  });

  it("parses crop and rotation; edits alt text and blip target", () => {
    const doc = loadWithImage();
    const para = doc.sections[0].blocks[0] as Paragraph;
    const run = para.children[0] as Run;
    const img = run.content.find((c) => c.kind === "image");
    if (!img || img.kind !== "image") throw new Error("no image");
    expect(img.crop?.l).toBeCloseTo(0.25, 3);
    expect(img.crop?.t).toBeCloseTo(0.1, 3);
    expect(img.rotation).toBeCloseTo(90, 1);
    // alt text
    expect(imageAltText(img.srcDrawing!)).toBe("old alt");
    expect(setImageAltText(doc, img.srcDrawing!, "new alt")).toBe(true);
    expect(imageAltText(img.srcDrawing!)).toBe("new alt");
    // replace blip
    const relId = doc.addImageResource(PNG, "png");
    expect(replaceImageBlip(doc, img.srcDrawing!, relId)).toBe(true);
    const para2 = doc.sections[0].blocks[0] as Paragraph;
    const img2 = (para2.children[0] as Run).content.find((c) => c.kind === "image");
    expect(img2 && img2.kind === "image" ? img2.part : "").toContain("media/image");
  });
});

describe("insertFootnote", () => {
  it("creates footnotes.xml on demand, splits at the caret, round-trips", () => {
    const doc = loadDoc(p("Citation needed here"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    const id = insertFootnote(doc, t, 15, "See Smith 2024.");
    expect(id).toBe(1);
    expect(doc.footnotes.get(1)).toBeTruthy();
    // reference is between "needed" and " here"
    const doc2 = DocxDocument.load(doc.save());
    expect(doc2.footnotes.get(1)).toBeTruthy();
  });
});

describe("footnote editing", () => {
  // First text content of footnote `id`, or null.
  const noteText = (doc: DocxDocument, id: number) => {
    const blocks = doc.footnotes.get(id);
    const para = blocks?.find((b) => b.type === "paragraph") as Paragraph | undefined;
    for (const c of para?.children ?? []) {
      for (const r of c.type === "run" ? [c] : c.runs) {
        const tc = r.content.find((rc) => rc.kind === "text");
        if (tc && tc.kind === "text") return tc;
      }
    }
    return null;
  };

  it("keeps footnote text source refs (editable) and exposes the part as an editable root", () => {
    const doc = loadDoc(p("Citation needed here"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    insertFootnote(doc, t, 15, "See Smith 2024.");
    const tc = noteText(doc, 1);
    expect(tc, "footnote has a text content").toBeTruthy();
    // The source w:t is KEPT (v1 stripped it → render-only).
    expect(tc!.srcT).toBeTruthy();
    // The footnotes part is now an editable root and reachable by findParentOf.
    const roots = doc.editableRoots();
    const noteInRoot = roots.some((root) => {
      const walk = (el: XmlElement): boolean =>
        el === tc!.srcT || el.children.some(walk);
      return walk(root);
    });
    expect(noteInRoot, "footnote t reachable from an editable root").toBe(true);
    expect(doc.findParentOf(tc!.srcT!)).toBeTruthy();
  });

  it("edits a footnote's text and round-trips through footnotes.xml on save", () => {
    const doc = loadDoc(p("Citation needed here"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    insertFootnote(doc, t, 15, "See Smith 2024.");
    const tc = noteText(doc, 1)!;
    // Simulate a caret edit: mutate the source w:t, flag the part, relayout.
    tc.srcT!.text = tc.srcT!.text.replace("Smith", "Smith & Jones");
    doc.markDirtyIfFootnote(tc.srcT!);
    doc.refresh();
    expect(noteText(doc, 1)!.text).toContain("Smith & Jones");
    // Persisted to footnotes.xml.
    const reloaded = DocxDocument.load(doc.save());
    expect(noteText(reloaded, 1)!.text).toContain("Smith & Jones");
  });

  it("records a suggesting-mode edit inside a footnote as a w:ins in footnotes.xml", async () => {
    const { insertSuggestedText } = await import("../src/edit/suggest.js");
    const doc = loadDoc(p("Citation needed here"));
    const t = firstRun(doc).run.content.find((c) => c.kind === "text")!.srcT!;
    insertFootnote(doc, t, 15, "See Smith 2024.");
    const tc = noteText(doc, 1)!;
    let revId = 500;
    // findParentOf must reach the footnote run for the suggest core to work.
    const nc = insertSuggestedText(doc, tc.srcT!, 3, "XYZ", {
      author: "Alex",
      date: "2026-07-13T00:00:00Z",
      nextId: () => revId++,
    });
    expect(nc, "suggested insert into a footnote resolves its run").toBeTruthy();
    doc.markDirtyIfFootnote(nc!.t);
    const fnXml = DocxDocument.load(doc.save()).pkg.text("word/footnotes.xml");
    expect(fnXml).toContain("XYZ");
    expect(fnXml).toMatch(/<w:ins\b/);
  });
});

describe("insertPageField", () => {
  it("splits at the caret and inserts a PAGE fldSimple that round-trips", () => {
    const doc = loadDoc(p("Footer text here"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(insertPageField(doc, t, 11, "page")).toBe(true);
    const doc2 = DocxDocument.load(doc.save());
    const para = doc2.sections[0].blocks[0] as Paragraph;
    const fields = para.children.flatMap((c) =>
      (c.type === "run" ? [c] : c.runs).flatMap((r) => r.content.filter((rc) => rc.kind === "field")),
    );
    expect(fields.length).toBe(1);
    expect((fields[0] as { instruction: string }).instruction).toContain("PAGE");
  });

  it("pageOfTotal inserts Page {PAGE} of {NUMPAGES}", () => {
    const doc = loadDoc(p("x"));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(insertPageField(doc, t, 1, "pageOfTotal")).toBe(true);
    const doc2 = DocxDocument.load(doc.save());
    const para = doc2.sections[0].blocks[0] as Paragraph;
    const runs = para.children.flatMap((c) => (c.type === "run" ? [c] : c.runs));
    const instrs = runs.flatMap((r) => r.content.filter((rc) => rc.kind === "field"))
      .map((f) => (f as { instruction: string }).instruction);
    expect(instrs.some((i) => i.includes("PAGE"))).toBe(true);
    expect(instrs.some((i) => i.includes("NUMPAGES"))).toBe(true);
    const text = runs.flatMap((r) => r.content).map((rc) => (rc.kind === "text" ? rc.text : "")).join("");
    expect(text).toContain("Page ");
    expect(text).toContain(" of ");
  });

  it("inserts generic and live date/time fields with Word field instructions", () => {
    const doc = loadDoc(p("Fields: "));
    const { run } = firstRun(doc);
    const t = run.content.find((c) => c.kind === "text")!.srcT!;
    expect(insertField(doc, t, t.text.length, "NUMPAGES \\* MERGEFORMAT", "1")).toBe(true);
    expect(insertDateTimeField(doc, t, t.text.length, "date", "MMMM d, yyyy")).toBe(true);
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toContain('w:instr=" NUMPAGES \\* MERGEFORMAT "');
    expect(xml).toContain('w:instr=" DATE \\@ &quot;MMMM d, yyyy&quot; \\* MERGEFORMAT "');
    expect(xml).toMatch(/<w:fldSimple[^>]*>[\s\S]*?<\/w:fldSimple><w:r><w:t xml:space="preserve"\/><\/w:r>/);

    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const instructions = para.children.flatMap((child) =>
      (child.type === "run" ? [child] : child.runs).flatMap((oneRun) =>
        oneRun.content.filter((content) => content.kind === "field").map((content) => content.kind === "field" ? content.instruction : ""),
      ),
    );
    expect(instructions.some((instruction) => instruction.includes("NUMPAGES"))).toBe(true);
    expect(instructions.some((instruction) => instruction.includes('DATE \\@ "MMMM d, yyyy"'))).toBe(true);
  });

  it("inserts a field at the exact text position inside a multi-content run", () => {
    const doc = loadDoc(`<w:p><w:r><w:t>BeforeAfter</w:t><w:tab/><w:t>Suffix</w:t></w:r></w:p>`);
    const { run } = firstRun(doc);
    const t = run.content.find((content) => content.kind === "text")!.srcT!;
    expect(insertField(doc, t, 6, "PAGE")).toBe(true);
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toMatch(/>Before<\/w:t><\/w:r><w:fldSimple[^>]*><w:r>[\s\S]*?<\/w:r><\/w:fldSimple><w:r><w:t[^>]*>After<\/w:t><w:tab\/><w:t>Suffix<\/w:t>/);
  });
});

describe("bookmarks and cross-references", () => {
  it("validates Word bookmark names", () => {
    expect(validBookmarkName("Quarterly_Revenue2")).toBe(true);
    expect(validBookmarkName("2Revenue")).toBe(false);
    expect(validBookmarkName("Quarterly Revenue")).toBe(false);
    expect(validBookmarkName("A".repeat(41))).toBe(false);
  });

  it("wraps a partial selection and preserves it through save-back", () => {
    const doc = loadDoc(p("Quarterly Revenue"));
    const { run } = firstRun(doc);
    expect(insertBookmarkAroundSelection(doc, [segFor(run, 0, 9)], "Quarterly")).toBe(true);
    expect(listBookmarks(doc)).toEqual(["Quarterly"]);
    expect(insertBookmarkAroundSelection(doc, [segFor(run, 0, 9)], "Quarterly")).toBe(false);
    const xml = DocxDocument.load(doc.save()).pkg.text("word/document.xml");
    expect(xml).toMatch(/<w:bookmarkStart[^>]+w:name="Quarterly"[^>]*><w:r>[\s\S]*?<w:t[^>]*>Quarterly<\/w:t><\/w:r><w:bookmarkEnd/);
  });

  it("inserts a zero-length bookmark at an exact caret offset", () => {
    const doc = loadDoc(p("HeadTail"));
    const { run } = firstRun(doc);
    const t = run.content.find((content) => content.kind === "text")!.srcT!;
    expect(insertBookmarkAt(doc, t, 4, "Middle")).toBe(true);
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toMatch(/>Head<\/w:t><\/w:r><w:bookmarkStart[^>]+w:name="Middle"[^>]*><w:bookmarkEnd[^>]*><w:r><w:t[^>]*>Tail<\/w:t>/);
  });

  it("inserts live REF and PAGEREF fields only for existing bookmarks", () => {
    const doc = loadDoc(p("Quarterly Revenue") + p("See "));
    const first = firstRun(doc).run;
    expect(insertBookmarkAroundSelection(doc, [segFor(first, 0, 17)], "Revenue")).toBe(true);
    const second = firstRun(doc, 1).run;
    const target = second.content.find((content) => content.kind === "text")!.srcT!;
    expect(insertCrossReference(doc, target, target.text.length, "Missing", "text")).toBe(false);
    expect(insertCrossReference(doc, target, target.text.length, "Revenue", "text")).toBe(true);
    expect(insertCrossReference(doc, target, target.text.length, "Revenue", "page")).toBe(true);
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toContain('w:instr=" REF Revenue \\h \\* MERGEFORMAT "');
    expect(xml).toContain('w:instr=" PAGEREF Revenue \\h \\* MERGEFORMAT "');
  });
});

describe("math editing", () => {
  it("linearizes n-ary, delimiter and matrix nodes", () => {
    expect(
      linearizeMath([
        { t: "nary", chr: "\u2211", sub: [{ t: "run", text: "i=1" }], sup: [{ t: "run", text: "n" }], e: [{ t: "run", text: "i" }] },
        { t: "dlm", beg: "(", end: ")", e: [[{ t: "run", text: "x" }]] },
        { t: "mat", rows: [[[{ t: "run", text: "a" }], [{ t: "run", text: "b" }]], [[{ t: "run", text: "c" }], [{ t: "run", text: "d" }]]] },
      ]),
    ).toBe("\u2211_{i=1}^ni(x)[a&b;c&d]");
  });

  it("linear form round-trips through the parser", () => {
    const cases = ["e^x=1+x+x/2", "a_i+b^{2y}", "{a+b}/{2c}", "√{x+1}", "√[3]{x+1}", "√[5]{x+1}"];
    for (const c of cases) {
      expect(linearizeMath(parseMathLinear(c))).toBe(c);
    }
    // ∛/∜ are accepted as input shorthands and canonicalize to √[n]{…} so the
    // degree stays editable as a plain digit.
    expect(linearizeMath(parseMathLinear("∛{x+1}"))).toBe("√[3]{x+1}");
    expect(linearizeMath(parseMathLinear("∜{x+1}"))).toBe("√[4]{x+1}");
  });

  it("rewrites an equation from linear text", () => {
    const XML = `<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
    </w:p>`;
    const doc = loadDoc(XML);
    const para = doc.sections[0].blocks[0] as Paragraph;
    const math = (para.children[0] as Run).content.find((c) => c.kind === "math");
    if (!math || math.kind !== "math" || !math.src) throw new Error();
    expect(setMathLinear(doc, math.src, "a^2+b^2=c^2")).toBe(true);
    const para2 = doc.sections[0].blocks[0] as Paragraph;
    const math2 = (para2.children[0] as Run).content.find((c) => c.kind === "math");
    if (!math2 || math2.kind !== "math") throw new Error();
    expect(mathLinearOf(doc, math2.src!)).toBe("a^2+b^2=c^2");
    expect(math2.nodes.filter((n) => n.t === "sup").length).toBe(3);
    // and it round-trips through save
    const doc3 = DocxDocument.load(doc.save());
    const para3 = doc3.sections[0].blocks[0] as Paragraph;
    const math3 = (para3.children[0] as Run).content.find((c) => c.kind === "math");
    expect(math3 && math3.kind === "math" ? mathLinearOf(doc3, math3.src!) : "").toBe("a^2+b^2=c^2");
  });

  it("inserts a new editable OMML equation at the caret", () => {
    const doc = loadDoc(p("BeforeAfter"));
    const { run } = firstRun(doc);
    const t = run.content.find((content) => content.kind === "text")!.srcT!;
    expect(insertMathAt(doc, t, 6, "x^2+y/2")).not.toBeNull();
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(xml).toMatch(/>Before<\/w:t><\/w:r><m:oMath>/);
    expect(xml).toMatch(/<\/m:oMath><w:r><w:t[^>]*>After<\/w:t>/);

    const reloaded = DocxDocument.load(doc.save());
    const para = reloaded.sections[0].blocks[0] as Paragraph;
    const math = para.children.flatMap((child) => child.type === "run" ? child.content : child.runs.flatMap((oneRun) => oneRun.content))
      .find((content) => content.kind === "math");
    expect(math && math.kind === "math" ? mathLinearOf(reloaded, math.src!) : "").toBe("x^2+y/2");
  });

  it("declares the math namespace on the edited header part", () => {
    const doc = loadDoc(
      `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>`,
      {
        "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
        "word/header1.xml": `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
      },
    );
    const header = doc.headers.get("rIdH");
    const run = header?.blocks[0]?.type === "paragraph" ? header.blocks[0].children[0] : null;
    if (!run || run.type !== "run") throw new Error("header run missing");
    const t = run.content.find((content) => content.kind === "text")?.srcT;
    if (!t) throw new Error("header text missing");
    expect(insertMathAt(doc, t, t.text.length, "x^2")).not.toBeNull();
    const saved = DocxDocument.load(doc.save()).pkg.text("word/header1.xml");
    expect(saved).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(saved).toContain("<m:oMath>");
  });
});

describe("line numbering (w:lnNumType)", () => {
  const SECT =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const lnDoc = () => loadDoc(p("body text") + SECT);
  const firstT = (doc: DocxDocument) => {
    const para = doc.sections[0].blocks[0] as Paragraph;
    return ((para.children[0] as Run).content[0] as TextContent).srcT!;
  };

  it("enables line numbering with countBy and restart, readable via lineNumberingAt", async () => {
    const { setLineNumbering, lineNumberingAt } = await import("../src/edit/sections.js");
    const doc = lnDoc();
    expect(lineNumberingAt(doc, firstT(doc))).toBeNull();
    expect(setLineNumbering(doc, { enabled: true, countBy: 5, restart: "continuous" })).toBe(true);
    expect(doc.sections[0].props.lineNumbering?.countBy).toBe(5);
    const ln = lineNumberingAt(doc, firstT(doc));
    expect(ln).toEqual({ countBy: 5, restart: "continuous", start: 1 });
  });

  it("newPage restart and start=1 stay out of the XML (OOXML defaults)", async () => {
    const { setLineNumbering } = await import("../src/edit/sections.js");
    const doc = lnDoc();
    setLineNumbering(doc, { enabled: true, countBy: 1, restart: "newPage", start: 1 });
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml).toContain("lnNumType");
    expect(xml).not.toContain('restart="newPage"');
    expect(xml).not.toContain("start=");
  });

  it("round-trips through save and can be turned back off", async () => {
    const { setLineNumbering, lineNumberingAt } = await import("../src/edit/sections.js");
    const doc = lnDoc();
    setLineNumbering(doc, { enabled: true, countBy: 10, restart: "newSection" });
    const reloaded = DocxDocument.load(doc.save());
    const rt = lineNumberingAt(reloaded, firstT(reloaded));
    expect(rt).toEqual({ countBy: 10, restart: "newSection", start: 1 });
    // Disabling removes the element entirely.
    expect(setLineNumbering(reloaded, { enabled: false })).toBe(true);
    expect(lineNumberingAt(reloaded, firstT(reloaded))).toBeNull();
    const off = DocxDocument.load(reloaded.save());
    expect(off.sections[0].props.lineNumbering).toBeUndefined();
  });

  it("inserts lnNumType in schema order (after pgMar, before cols)", async () => {
    const { setLineNumbering } = await import("../src/edit/sections.js");
    const { setPageLayout } = await import("../src/edit/blocks.js");
    const doc = lnDoc();
    setPageLayout(doc, { columns: 2 }); // adds <w:cols> after pgMar
    setLineNumbering(doc, { enabled: true, countBy: 1 });
    const xml = serializeXml(doc.editableRoots()[0]);
    expect(xml.indexOf("pgMar")).toBeLessThan(xml.indexOf("lnNumType"));
    expect(xml.indexOf("lnNumType")).toBeLessThan(xml.indexOf("<w:cols"));  });
});

describe("math editing safety", () => {
  const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
  const load = (omml: string) => {
    const doc = loadDoc(`<w:p ${M}>${omml}</w:p>`);
    const para = doc.sections[0].blocks[0] as Paragraph;
    const math = (para.children[0] as Run).content.find((c) => c.kind === "math");
    if (!math || math.kind !== "math" || !math.src) throw new Error("no math");
    return { doc, para, srcOf: () => (((doc.sections[0].blocks[0] as Paragraph).children[0] as Run).content.find((c) => c.kind === "math") as { src: XmlElement }).src };
  };
  const run = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`;
  const INTEGRAL = `<m:oMath><m:nary><m:naryPr><m:chr m:val="∫"/></m:naryPr><m:sub>${run("a")}</m:sub><m:sup>${run("b")}</m:sup><m:e>${run("x")}</m:e></m:nary></m:oMath>`;
  const DELIM = `<m:oMath><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${run("x")}</m:e></m:d></m:oMath>`;
  const MATRIX = `<m:oMath><m:m><m:mr><m:e>${run("a")}</m:e><m:e>${run("b")}</m:e></m:mr><m:mr><m:e>${run("c")}</m:e><m:e>${run("d")}</m:e></m:mr></m:m></m:oMath>`;
  const ACCENT = `<m:oMath><m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>${run("x")}</m:e></m:acc></m:oMath>`;
  const LIMIT = `<m:oMath><m:limLow><m:e>${run("lim")}</m:e><m:lim>${run("n")}</m:lim></m:limLow></m:oMath>`;

  it("classifies round-trippable equations as editable", () => {
    for (const omml of [INTEGRAL, DELIM, MATRIX]) {
      const { srcOf } = load(omml);
      expect(isLinearSafe(srcOf())).toBe(true);
    }
  });

  it("classifies structure-only equations (accent, limit) as read-only", () => {
    for (const omml of [ACCENT, LIMIT]) {
      const { srcOf } = load(omml);
      expect(isLinearSafe(srcOf())).toBe(false);
    }
  });

  it("linear parser reconstructs n-ary, delimiter and matrix structurally", () => {
    expect(parseMathLinear("∫_a^bx")[0].t).toBe("nary");
    expect(parseMathLinear("(x)")[0].t).toBe("dlm");
    const mat = parseMathLinear("[a&b;c&d]")[0];
    expect(mat.t).toBe("mat");
    if (mat.t === "mat") expect(mat.rows.length).toBe(2);
  });

  it("editing a matrix cell keeps it a matrix (no collapse to literal text)", () => {
    const { doc, srcOf } = load(MATRIX);
    const linear = mathLinearOf(doc, srcOf());
    expect(linear).toBe("[a&b;c&d]");
    expect(setMathLinear(doc, srcOf(), "[a&b;c&e]")).toBe(true);
    const after = srcOf();
    expect(mathLinearOf(doc, after)).toBe("[a&b;c&e]");
    // still a real m:m element, not a run of literal "[a&b;c&e]"
    const hasMatrix = (e: XmlElement): boolean =>
      e.name.endsWith("m") && e.children.some((c) => c.name.endsWith("mr")) ? true : e.children.some(hasMatrix);
    expect(hasMatrix(after)).toBe(true);
  });

  it("radical degree (nth root) survives linearization and editing", () => {
    const CUBE = `<m:oMath><m:rad><m:deg>${run("3")}</m:deg><m:e>${run("x+1")}</m:e></m:rad></m:oMath>`;
    const { doc, srcOf } = load(CUBE);
    expect(mathLinearOf(doc, srcOf())).toBe("√[3]{x+1}");
    expect(isLinearSafe(srcOf())).toBe(true);
    expect(setMathLinear(doc, srcOf(), "√[5]{y+2}")).toBe(true);
    const after = srcOf();
    expect(mathLinearOf(doc, after)).toBe("√[5]{y+2}");
    const degText = (e: XmlElement, inDeg = false): string => {
      const here = inDeg && localName(e.name) === "t" ? e.text : "";
      return here + e.children.map((c) => degText(c, inDeg || localName(e.name) === "deg")).join("");
    };
    expect(degText(after)).toBe("5");
  });

  it("a hidden degree (m:degHide) stays a plain square root", () => {
    const HIDDEN =
      `<m:oMath><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>` +
      `<m:deg></m:deg><m:e>${run("x")}</m:e></m:rad></m:oMath>`;
    const { doc, srcOf } = load(HIDDEN);
    expect(mathLinearOf(doc, srcOf())).toBe("√x");
    expect(isLinearSafe(srcOf())).toBe(true);
  });

  it("n-ary integrand survives a round-trip through OMML", () => {
    const { doc, srcOf } = load(INTEGRAL);
    expect(mathLinearOf(doc, srcOf())).toBe("∫_a^bx");
    expect(setMathLinear(doc, srcOf(), "∫_a^by")).toBe(true);
    const after = srcOf();
    expect(mathLinearOf(doc, after)).toBe("∫_a^by");
    const isNary = (e: XmlElement): boolean =>
      e.name.endsWith("nary") ? true : e.children.some(isNary);
    expect(isNary(after)).toBe(true);
  });
});

describe("selectionTextLogical (clipboard copy order)", () => {
  it("emits logical order when bidi segments arrive reversed (visual paint order)", () => {
    // A bidi/RTL paragraph's runs paint right-to-left, so getSelectionSegments
    // hands them back reversed; copy must still read in source order.
    const doc = loadDoc(
      `<w:p>` +
        `<w:r><w:t xml:space="preserve">שלום</w:t></w:r>` +
        `<w:r><w:t xml:space="preserve"> </w:t></w:r>` +
        `<w:r><w:t xml:space="preserve">עולם</w:t></w:r>` +
        `</w:p>`,
    );
    const runs = (doc.sections[0].blocks[0] as Paragraph).children.filter(
      (c) => c.type === "run",
    ) as Run[];
    const logical = [segFor(runs[0], 0, 4), segFor(runs[1], 0, 1), segFor(runs[2], 0, 4)];
    const visual = [...logical].reverse();
    expect(selectionTextLogical(doc, visual)).toBe("שלום עולם");
    // Order-independent: same result no matter how the caller ordered them.
    expect(selectionTextLogical(doc, logical)).toBe("שלום עולם");
  });

  it("keeps LTR single-run order and slices offsets", () => {
    const doc = loadDoc(p("Hello world"));
    const { run } = firstRun(doc);
    expect(selectionTextLogical(doc, [segFor(run, 0, 11)])).toBe("Hello world");
    expect(selectionTextLogical(doc, [segFor(run, 6, 11)])).toBe("world");
  });

  it("joins paragraphs with a newline in document order", () => {
    const doc = loadDoc(p("First") + p("Second"));
    const b0 = doc.sections[0].blocks[0] as Paragraph;
    const b1 = doc.sections[0].blocks[1] as Paragraph;
    const s0 = segFor(b0.children[0] as Run, 0, 5);
    const s1 = segFor(b1.children[0] as Run, 0, 6);
    expect(selectionTextLogical(doc, [s0, s1])).toBe("First\nSecond");
  });

  it("returns empty string for an empty selection", () => {
    const doc = loadDoc(p("x"));
    expect(selectionTextLogical(doc, [])).toBe("");
  });
});
