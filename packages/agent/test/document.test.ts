import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { DocxDocument, localName, type XmlElement } from "@wordinweb/core";
import { INTENT_KINDS } from "@wordinweb/collab/client";
import { AgentDocument, validateAgentOperationShape } from "../src/index.js";
import { anchorBox, body, makeDocx, makeDocxWithHeader } from "./helpers.js";

const PNG = Uint8Array.from(Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC",
  "base64",
));

function textOf(doc: DocxDocument): string {
  const collect = (element: XmlElement): string =>
    (localName(element.name) === "t" ? element.text : "") + element.children.map(collect).join("");
  return collect(doc.docRoot);
}

describe("AgentDocument local tools", () => {
  it("loads, inspects, edits, and saves without a DOM", async () => {
    const agent = AgentDocument.create({ provenance: { author: "Test Agent", now: () => "2026-07-31T00:00:00Z", nextId: () => "ABCDEF01" } });
    const overview = agent.inspect({ kind: "overview" });
    expect(overview).toMatchObject({ revision: "0", sections: 1, blocks: { paragraphs: 1, tables: 0 } });

    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    const run = paragraph.runs[0];
    const result = await agent.edit({
      revision: "0",
      operations: [{ kind: "insertText", at: { blockRef: paragraph.ref, runRef: run.ref, offset: 0 }, text: "Hello agent" }],
    });
    expect(result).toMatchObject({ revision: "1", status: "applied", connection: "local" });
    expect(textOf(DocxDocument.load(agent.save()))).toContain("Hello agent");
  });

  it("creates native heading structure in a blank document", async () => {
    const agent = AgentDocument.create();
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    await agent.edit({
      revision: agent.revision,
      operations: [
        { kind: "insertText", at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 0 }, text: "Architecture Options" },
        { kind: "formatParagraph", blockRef: paragraph.ref, styleId: "Heading1" },
      ],
    });
    const overview = agent.inspect({ kind: "overview" });
    expect("outline" in overview && overview.outline).toEqual([
      expect.objectContaining({ ref: paragraph.ref, level: 1, text: "Architecture Options", styleId: "Heading1" }),
    ]);
    expect(DocxDocument.load(agent.save()).styles.byId.has("Heading1")).toBe(true);
  });

  it("authors a whole table, content and header included, in one insertTable operation", async () => {
    const agent = AgentDocument.load(makeDocx(body(`<w:p><w:r><w:t>intro</w:t></w:r></w:p>`)));
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    await agent.edit({
      revision: agent.revision,
      operations: [{
        kind: "insertTable",
        runRef: paragraph.runs[0].ref,
        rows: 2,
        cols: 3,
        cells: [["Quarter", "Region", "Sales"], ["Q1", "West", "100"]],
        headerRow: true,
      }],
    });
    const overview = agent.inspect({ kind: "overview" });
    expect(overview).toMatchObject({ blocks: { tables: 1 } });
    const saved = textOf(DocxDocument.load(agent.save()));
    for (const text of ["Quarter", "Region", "Sales", "Q1", "West", "100"]) expect(saved).toContain(text);
    expect(validateAgentOperationShape({ kind: "insertTable", runRef: "run:1", rows: 2, cols: 2, cells: [[1]] })).toContain("cells");
  });

  it("advertises the compose tool only where compose can succeed", () => {
    const created = AgentDocument.create().tools().map((tool) => tool.name);
    expect(created).toContain("word_document_compose");
    const loaded = AgentDocument.load(makeDocx(body(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`)));
    expect(loaded.tools().map((tool) => tool.name)).not.toContain("word_document_compose");
  });

  it("composes a rich new document in one schema-enforced request", async () => {
    const agent = AgentDocument.create();
    const logo = agent.addAsset(PNG, "image/png");
    const result = await agent.compose({
      revision: agent.revision,
      page: { margins: { top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 } },
      body: [
        { type: "paragraph", text: "Systems Decision Brief", styleId: "Title" },
        { type: "paragraph", text: "Architecture review", styleId: "Subtitle" },
        { type: "heading", level: 1, text: "Context" },
        { type: "paragraph", runs: [{ text: "Decision: ", bold: true }, { text: "select the platform." }], spacing: { afterPt: 8 } },
        { type: "table", headerRows: 1, rows: [["Option", "Score"], ["Managed", "9"], ["Custom", "7"]] },
        { type: "equation", mathText: "x=(-b+-sqrt(b^2-4ac))/(2a)", align: "center" },
        { type: "chart", chart: { type: "column", title: "Decision profile", categories: ["Control", "Speed"], series: [{ name: "Managed", values: [8, 9] }, { name: "Custom", values: [10, 6] }] }, widthPx: 520, heightPx: 280, align: "center" },
        { type: "smartArt", smartArt: { layout: "process", items: ["Discover", "Assess", "Decide", "Deliver"] }, widthPx: 520, heightPx: 220, align: "center" },
        { type: "image", assetRef: logo, widthPx: 96, heightPx: 24, alt: "WordInWeb mark", wrap: "square", position: { xPx: 620, yPx: 80 } },
        { type: "shape", preset: "rectangle", widthPx: 180, heightPx: 54, position: { xPx: 80, yPx: 700 }, fill: "D9EAF7", line: { color: "1F4E79", widthPx: 2, dash: "solid" }, text: "Decision gate" },
        { type: "pageBreak" },
        { type: "heading", level: 1, text: "Recommendation" },
        { type: "paragraph", text: "Adopt the managed platform." },
      ],
      header: [
        { type: "image", assetRef: logo, widthPx: 96, heightPx: 24, alt: "Header mark" },
        { type: "wordArt", text: "DRAFT", preset: "plain", widthPx: 420, heightPx: 100, position: { xPx: 190, yPx: 420 }, rotation: 315, fill: "B7C9DB", opacity: 0.12, wrap: "behind", order: "back" },
      ],
      footer: [{ type: "pageNumber", fieldKind: "pageOfTotal", align: "center", color: "627D98", fontSizePt: 8, fontFamily: "Arial" }],
    });
    expect(result).toMatchObject({ revision: "1", status: "applied", blocks: 16, components: 8 });

    const overview = agent.inspect({ kind: "overview" });
    expect(overview).toMatchObject({ sections: 1, blocks: { tables: 1 } });
    expect("outline" in overview && overview.outline.map((entry) => entry.text)).toEqual(["Context", "Recommendation"]);
    expect("stories" in overview && overview.stories.map((story) => story.kind)).toEqual(expect.arrayContaining(["body", "header", "footer"]));

    const bodyRead = agent.inspect({ kind: "read", story: "body", maxBlocks: 100, maxCharacters: 100_000 });
    if (!("blocks" in bodyRead)) throw new Error("composed body missing");
    const bodyTypes = bodyRead.blocks.flatMap((block) => block.type === "paragraph" ? block.runs.flatMap((run) => run.components.map((component) => component.type)) : []);
    expect(bodyTypes).toEqual(expect.arrayContaining(["math", "chart", "smart-art", "anchored-image", "anchored-textbox"]));

    const headerStory = "stories" in overview ? overview.stories.find((story) => story.kind === "header") : undefined;
    const footerStory = "stories" in overview ? overview.stories.find((story) => story.kind === "footer") : undefined;
    if (!headerStory || !footerStory) throw new Error("composed repeating stories missing");
    const headerRead = agent.inspect({ kind: "read", story: headerStory.id });
    const footerRead = agent.inspect({ kind: "read", story: footerStory.id });
    if (!("blocks" in headerRead) || !("blocks" in footerRead)) throw new Error("composed story content missing");
    const headerTypes = headerRead.blocks.flatMap((block) => block.type === "paragraph" ? block.runs.flatMap((run) => run.components.map((component) => component.type)) : []);
    const footerTypes = footerRead.blocks.flatMap((block) => block.type === "paragraph" ? block.runs.flatMap((run) => run.components.map((component) => component.type)) : []);
    expect(headerTypes).toEqual(expect.arrayContaining(["image", "anchored-textbox"]));
    expect(footerTypes).toContain("field");
    const watermark = headerRead.blocks.flatMap((block) => block.type === "paragraph" ? block.runs.flatMap((run) => run.components) : [])
      .find((component) => component.type === "anchored-textbox");
    if (!watermark) throw new Error("composed watermark missing");
    expect(agent.inspect({ kind: "object", ref: watermark.ref })).toMatchObject({ detail: { wordArt: true, fill: "#B7C9DB", opacity: 0.12 } });

    const reopened = AgentDocument.load(agent.save());
    expect(reopened.inspect({ kind: "overview" })).toMatchObject({ blocks: { tables: 1 } });
    const pkg = DocxDocument.load(agent.save()).pkg;
    const headerRels = pkg.text("word/_rels/header1.xml.rels");
    const documentRels = pkg.text("word/_rels/document.xml.rels");
    expect(headerRels.match(/\/relationships\/image/g)).toHaveLength(1);
    expect(documentRels.match(/\/relationships\/image/g)).toHaveLength(1);
    expect(headerRels.match(/Target="([^"]+\.png)"/)?.[1]).not.toBe(documentRels.match(/Target="([^"]+\.png)"/)?.[1]);
    expect(pkg.text("word/header1.xml")).toContain('w14:alpha w14:val="12000"');
    expect(pkg.text("word/footer1.xml")).toContain('<w:jc w:val="center"/>');
    expect(pkg.text("word/footer1.xml")).toContain('<w:color w:val="627D98"/>');
    expect(pkg.text("word/document.xml")).toContain('<wp:extent cx="4953000" cy="2667000"/>');
    expect(pkg.text("word/document.xml")).toContain('<wp:extent cx="4953000" cy="2095500"/>');
    await expect(agent.compose({ revision: agent.revision, body: [{ type: "paragraph", text: "replace" }] })).rejects.toThrow("new AgentDocument.create");
  });

  it("keeps a new document unchanged when composition validation fails", async () => {
    const agent = AgentDocument.create();
    const before = agent.save();
    await expect(agent.compose({
      revision: agent.revision,
      body: [{ type: "table", rows: [["A", "B"], ["C"]] }],
    })).rejects.toThrow("equal lengths");
    await expect(agent.compose({
      revision: agent.revision,
      body: [{ type: "chart", chart: { type: "column", categories: ["A"], series: [{ name: "S", values: [1] }] }, widthPx: 400 }],
    })).rejects.toThrow("widthPx and heightPx");
    await expect(agent.compose({
      revision: agent.revision,
      body: [{ type: "wordArt", text: "DRAFT", preset: "plain", fill: "AABBCC", opacity: 2 }],
    })).rejects.toThrow("opacity");
    expect(agent.revision).toBe("0");
    expect(agent.save()).toEqual(before);
    await expect(AgentDocument.load(before).compose({ revision: "0", body: [{ type: "paragraph", text: "replace" }] }))
      .rejects.toThrow("new AgentDocument.create");
  });

  it("composes continuous lists, borderless shapes, first-page furniture, and compact object summaries", async () => {
    const agent = AgentDocument.create();
    const result = await agent.compose({
      revision: agent.revision,
      body: [
        { type: "paragraph", text: "One", list: "number" },
        { type: "paragraph", text: "Two", list: "number" },
        { type: "paragraph", text: "Three", list: "number" },
        {
          type: "shape",
          preset: "rectangle",
          text: "Cover panel",
          textStyle: { color: "F7F0E3", fontSizePt: 10, fontFamily: "Aptos", bold: true, align: "right" },
          widthPx: 400,
          heightPx: 300,
          fill: "10243E",
          line: null,
        },
      ],
      footer: [{ type: "pageNumber", fieldKind: "pageOfTotal", align: "center", color: "10243E" }],
      firstFooter: [{ type: "pageNumber", fieldKind: "pageOfTotal", align: "center", color: "FFFFFF" }],
    });

    expect(result.overview).toMatchObject({ objectCounts: expect.objectContaining({ "anchored-textbox": 1, field: 4 }) });
    expect(result.createdObjects).toEqual(expect.arrayContaining([
      expect.objectContaining({ type: "anchored-textbox", label: "Cover panel", editRef: expect.stringMatching(/^object:/) }),
      expect.objectContaining({ type: "field", label: expect.stringContaining("PAGE") }),
    ]));
    expect(result.createdObjectsTruncated).toBe(false);

    const pkg = DocxDocument.load(agent.save()).pkg;
    const documentXml = pkg.text("word/document.xml");
    const numIds = [...documentXml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((match) => match[1]);
    expect(numIds.slice(0, 3)).toEqual(["1", "1", "1"]);
    expect(pkg.text("word/numbering.xml")).toContain('<w:lvlText w:val="%1."/>');
    expect(documentXml).toContain('<w:footerReference w:type="first" r:id="rId5"/>');
    expect(documentXml).toContain("<w:titlePg/>");
    expect(pkg.text("word/footer2.xml")).toContain('<w:color w:val="FFFFFF"/>');
    expect(documentXml).toContain("<a:noFill/>");
    expect(documentXml).toContain('<w:color w:val="F7F0E3"/>');
    expect(documentXml).toContain('<w:sz w:val="20"/>');
    expect(documentXml).toContain('<w:rFonts w:ascii="Aptos"');
    expect(documentXml).toContain('<w:jc w:val="right"/>');
    expect(documentXml).toContain("<wp:wrapTopAndBottom/>");
  });

  it("rejects stale revisions and internal identifiers without mutating", async () => {
    const agent = AgentDocument.create();
    const before = agent.save();
    await expect(agent.edit({ revision: "stale", operations: [{ kind: "acceptAllRevisions" }] })).rejects.toThrow("stale");
    await expect(agent.edit({ revision: "0", operations: [{ kind: "insertText", at: { blockRef: "block:1", runRef: "run:2", offset: 0 }, text: "x", runId: 2 }] })).rejects.toThrow("internal");
    await expect(agent.edit({ revision: "0", operations: [{ kind: "insertText", at: { blockRef: "block:1", runRef: "run:2", offset: 0 }, text: "x", surprise: true }] })).rejects.toThrow("does not allow surprise");
    expect(agent.save()).toEqual(before);
    expect(agent.revision).toBe("0");
  });

  it("paginates a paragraph that is larger than the context budget", async () => {
    const agent = AgentDocument.create();
    const first = agent.inspect({ kind: "read" });
    if (!("blocks" in first) || first.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = first.blocks[0];
    await agent.edit({ revision: "0", operations: [{ kind: "insertText", at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 0 }, text: "x".repeat(25_000) }] });
    const page1 = agent.inspect({ kind: "read", maxCharacters: 10_000 });
    if (!("blocks" in page1)) throw new Error("missing read result");
    expect(page1.truncated).toBe(true);
    expect(page1.blocks[0].type === "paragraph" && page1.blocks[0].text.length).toBe(10_000);
    const page2 = agent.inspect({ kind: "read", cursor: page1.next, maxCharacters: 10_000 });
    if (!("blocks" in page2)) throw new Error("missing read result");
    expect(page2.blocks[0].type === "paragraph" && page2.blocks[0].textRange.start).toBe(10_000);
  });

  it("returns compact editable text from multiple stories in one inspection", async () => {
    const agent = AgentDocument.create();
    await agent.compose({
      revision: agent.revision,
      body: [
        { type: "paragraph", runs: [{ text: "Decision: ", bold: true }, { text: "select the platform." }] },
        { type: "paragraph", text: "" },
      ],
      header: [{ type: "paragraph", text: "Confidential" }],
      footer: [{ type: "paragraph", text: "Review copy" }],
    });

    const context = agent.inspect({ kind: "context" });
    if (!("contents" in context)) throw new Error("compact context missing");
    expect(context.contents.map((content) => content.kind)).toEqual(expect.arrayContaining(["body", "header", "footer"]));
    const bodyContext = context.contents.find((content) => content.story === "body");
    const paragraph = bodyContext?.blocks.find((block) => block.type === "paragraph");
    if (!paragraph || paragraph.type !== "paragraph") throw new Error("compact paragraph missing");
    expect(paragraph).toMatchObject({
      text: "Decision: select the platform.",
      runs: [
        { ref: expect.stringMatching(/^run:/), start: 0, end: 10 },
        { ref: expect.stringMatching(/^run:/), start: 10, end: 30 },
      ],
    });
    expect(bodyContext?.blocks.filter((block) => block.type === "paragraph")).toHaveLength(1);
    expect(paragraph).not.toHaveProperty("range");
    expect(paragraph).not.toHaveProperty("bookmarks");
    expect(paragraph).not.toHaveProperty("objects");
    expect(paragraph.runs[0]).not.toHaveProperty("formatting");

    const bodyOnly = agent.inspect({ kind: "context", stories: ["body"] });
    if (!("contents" in bodyOnly)) throw new Error("filtered context missing");
    expect(bodyOnly.contents.map((content) => content.story)).toEqual(["body"]);

    const overview = agent.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview missing");
    const detailed = overview.stories.map((story) => agent.inspect({ kind: "read", story: story.id, maxBlocks: 200, maxCharacters: 100_000 }));
    expect(JSON.stringify(context).length).toBeLessThan(JSON.stringify(detailed).length / 2);

    const page = agent.inspect({ kind: "context", maxCharacters: 12 });
    if (!("contents" in page)) throw new Error("bounded context missing");
    expect(page).toMatchObject({ truncated: true, remainingStories: expect.any(Array) });
    expect(page.contents[0].next).toBeTruthy();
    expect(page.contents[0].blocks[0]).toMatchObject({ text: "Decision: se", range: { start: 0, end: 12, total: 30 } });
  });

  it("returns detailed shape stories and deterministic overlap geometry", () => {
    const boxes = anchorBox(1, 457200, 457200, "AA0004") + anchorBox(2, 685800, 685800, "AA0005");
    const agent = AgentDocument.load(makeDocx(body(`<w:p>${boxes}<w:r><w:t>host</w:t></w:r></w:p>`)));
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const shapeRefs = read.blocks[0].runs.flatMap((run) => run.components).filter((component) => component.type === "anchored-textbox").map((component) => component.ref);
    expect(shapeRefs).toHaveLength(2);
    const context = agent.inspect({ kind: "context", include: ["objects"] });
    if (!("contents" in context)) throw new Error("compact context missing");
    expect(context.contents[0].blocks[0]).toMatchObject({ objects: [{ type: "anchored-textbox" }, { type: "anchored-textbox" }] });
    const detail = agent.inspect({ kind: "object", ref: shapeRefs[0] });
    expect(detail).toMatchObject({ type: "anchored-textbox" });
    expect(JSON.stringify(detail)).not.toContain("srcDrawing");

    const spatial = agent.inspect({ kind: "spatial", pages: { start: 1, count: 1 }, includeOverlaps: true });
    if (!("overlaps" in spatial)) throw new Error("missing spatial result");
    expect(spatial.objects.filter((object) => object.type === "anchored-textbox")).toHaveLength(2);
    expect(spatial.overlaps).toHaveLength(1);
    expect(spatial.overlaps[0].overlapArea).toBeGreaterThan(0);
    expect(spatial.overlaps[0].topObject).toBeTruthy();
  });

  it("edits a Word-authored VML line through its carrier reference", async () => {
    const documentXml =
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>` +
      `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
      `</w:sectPr></w:body></w:document>`;
    const headerXml =
      `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict>` +
      `<v:shape id="Numbers" style="position:absolute;margin-left:12pt;margin-top:18pt;width:24pt;height:600pt" ` +
      `fillcolor="#FFFFFF" strokecolor="#000000"><v:textbox><w:txbxContent>` +
      `<w:p><w:r><w:t>1</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>` +
      `<v:line id="Rule" style="position:absolute;z-index:3;mso-position-horizontal-relative:page;` +
      `mso-position-vertical-relative:page" from="48pt,24pt" to="48pt,624pt" ` +
      `strokecolor="#AA0000" strokeweight="1pt"><v:stroke dashstyle="solid"/></v:line>` +
      `</w:pict></w:r></w:p></w:hdr>`;
    const agent = AgentDocument.load(makeDocxWithHeader(documentXml, headerXml));
    const overview = agent.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview missing");
    const header = overview.stories.find((story) => story.kind === "header");
    if (!header) throw new Error("header missing");
    const read = agent.inspect({ kind: "read", story: header.id });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("line carrier missing");
    const component = read.blocks[0].runs.flatMap((run) => run.components)
      .find((candidate) => candidate.type === "anchored-line");
    if (!component) throw new Error("line object missing");
    const object = agent.inspect({ kind: "object", ref: component.ref });
    if (!("editRef" in object) || !object.editRef) throw new Error("line edit reference missing");
    expect(object.editRef).toMatch(/^object:[0-9]+:1$/);
    expect(object).toMatchObject({
      type: "anchored-line",
      detail: { x1: 64, x2: 64, color: "#AA0000", style: "single" },
    });

    await agent.edit({
      revision: agent.revision,
      operations: [
        { kind: "setFloatingPagePosition", objectRef: object.editRef, xPx: 120, yPx: 80 },
        { kind: "resizeDrawing", objectRef: object.editRef, widthPx: 1, heightPx: 500 },
        { kind: "setDrawingLineStyle", objectRef: object.editRef, color: "156082", widthPx: 3, dash: "dotted" },
      ],
    });

    const saved = DocxDocument.load(agent.save()).pkg.text("word/header1.xml");
    expect(saved).toContain('id="Numbers"');
    expect(saved).toContain('margin-left:12pt');
    expect(saved).toContain('strokecolor="#000000"');
    expect(saved).toContain('from="90pt,60pt"');
    expect(saved).toContain('to="90pt,435pt"');
    expect(saved).toContain('strokecolor="156082"');
    expect(saved).toContain('dashstyle="dot"');
  });

  it("inspects and edits a text-free VML page background", async () => {
    const agent = AgentDocument.load(makeDocx(
      `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:pict>` +
      `<v:rect id="Cover" style="position:absolute;margin-left:0;margin-top:0;width:612pt;height:792pt;` +
      `mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:-2" ` +
      `fillcolor="#102A43" stroked="f"/>` +
      `</w:pict></w:r><w:r><w:t>Cover title</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    ));

    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("cover paragraph missing");
    const component = read.blocks[0].runs.flatMap((run) => run.components)
      .find((candidate) => candidate.type === "anchored-textbox");
    if (!component) throw new Error("cover background missing");
    const object = agent.inspect({ kind: "object", ref: component.ref });
    expect(object).toMatchObject({
      editRef: component.ref,
      type: "anchored-textbox",
      detail: { fill: "#102A43", width: 816, height: 1056, behind: true, zOrder: -2 },
    });

    const spatial = agent.inspect({ kind: "spatial", pages: { start: 1, count: 1 } });
    expect(spatial.objects).toContainEqual(expect.objectContaining({
      editRef: component.ref,
      type: "anchored-textbox",
      layer: "behind-text",
      bounds: expect.objectContaining({ width: 816, height: 1056 }),
    }));

    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "setDrawingFill", objectRef: component.ref, color: "1F4E79" }],
    });
    expect(DocxDocument.load(agent.save()).pkg.text("word/document.xml")).toContain('fillcolor="1F4E79"');
  });

  it("inserts a persistent image into a header-owned relationship", async () => {
    const agent = AgentDocument.create();
    await agent.edit({ revision: agent.revision, operations: [{ kind: "ensureHeaderFooter", hfKind: "header" }] });
    const overview = agent.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview missing");
    const header = overview.stories.find((story) => story.kind === "header");
    if (!header) throw new Error("header missing");
    const read = agent.inspect({ kind: "read", story: header.id });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("header paragraph missing");

    const assetRef = agent.addAsset(PNG, "image/png");
    await agent.edit({
      revision: agent.revision,
      operations: [{
        kind: "insertImage",
        runRef: read.blocks[0].runs[0].ref,
        assetRef,
        widthPx: 96,
        heightPx: 24,
      }],
    });

    const updated = agent.inspect({ kind: "read", story: header.id });
    if (!("blocks" in updated) || updated.blocks[0].type !== "paragraph") throw new Error("updated header missing");
    const component = updated.blocks[0].runs.flatMap((run) => run.components)
      .find((candidate) => candidate.type === "image");
    if (!component) throw new Error("header image missing");
    const object = agent.inspect({ kind: "object", ref: component.ref });
    expect(object).toMatchObject({ type: "image", detail: { width: 96, height: 24 } });

    const saved = DocxDocument.load(agent.save());
    expect(saved.pkg.text("word/_rels/header1.xml.rels")).toContain("/relationships/image");
    expect(saved.pkg.text("word/_rels/document.xml.rels")).not.toContain("/relationships/image");
    const reopened = AgentDocument.load(saved.save());
    const reopenedOverview = reopened.inspect({ kind: "overview" });
    if (!("stories" in reopenedOverview)) throw new Error("reopened overview missing");
    const reopenedHeader = reopenedOverview.stories.find((story) => story.kind === "header");
    if (!reopenedHeader) throw new Error("reopened header missing");
    const reopenedRead = reopened.inspect({ kind: "read", story: reopenedHeader.id });
    if (!("blocks" in reopenedRead) || reopenedRead.blocks[0].type !== "paragraph") throw new Error("reopened header content missing");
    expect(reopenedRead.blocks[0].runs.flatMap((run) => run.components))
      .toContainEqual(expect.objectContaining({ type: "image" }));
  });

  it("creates one editable vertical rule that repeats through a header story", async () => {
    const agent = AgentDocument.create();
    const bodyRead = agent.inspect({ kind: "read", story: "body" });
    if (!("blocks" in bodyRead) || bodyRead.blocks[0].type !== "paragraph") throw new Error("body paragraph missing");
    const bodyRun = bodyRead.blocks[0].runs[0];
    await agent.edit({ revision: agent.revision, operations: [{ kind: "ensureHeaderFooter", hfKind: "header" }] });

    const overview = agent.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview missing");
    const header = overview.stories.find((story) => story.kind === "header");
    if (!header) throw new Error("header missing");
    const headerRead = agent.inspect({ kind: "read", story: header.id });
    if (!("blocks" in headerRead) || headerRead.blocks[0].type !== "paragraph") throw new Error("header paragraph missing");
    const headerRun = headerRead.blocks[0].runs[0];

    await agent.edit({
      revision: agent.revision,
      operations: [
        { kind: "insertShape", runRef: headerRun.ref, preset: "verticalLine", text: "" },
        { kind: "insertBreak", runRef: bodyRun.ref, breakKind: "page" },
      ],
    });

    const updatedHeader = agent.inspect({ kind: "read", story: header.id });
    if (!("blocks" in updatedHeader) || updatedHeader.blocks[0].type !== "paragraph") throw new Error("updated header missing");
    const component = updatedHeader.blocks[0].runs.flatMap((run) => run.components)
      .find((candidate) => candidate.type === "anchored-art");
    if (!component) throw new Error("inserted line missing");
    const object = agent.inspect({ kind: "object", ref: component.ref });
    if (!("editRef" in object) || !object.editRef) throw new Error("inserted line edit reference missing");
    await agent.edit({
      revision: agent.revision,
      operations: [
        { kind: "setFloatingPagePosition", objectRef: object.editRef, xPx: 72, yPx: 96 },
        { kind: "resizeDrawing", objectRef: object.editRef, widthPx: 3, heightPx: 800 },
        { kind: "setDrawingLineStyle", objectRef: object.editRef, color: "156082", widthPx: 2, dash: "solid" },
      ],
    });

    const spatial = agent.inspect({ kind: "spatial", pages: { start: 1, count: 10 } });
    if (!("objects" in spatial)) throw new Error("spatial result missing");
    expect(spatial.totalPages).toBeGreaterThanOrEqual(2);
    const occurrences = spatial.objects.filter((candidate) => candidate.editRef === object.editRef);
    expect(occurrences).toHaveLength(spatial.totalPages);
    expect(new Set(occurrences.map((candidate) => candidate.page)).size).toBe(spatial.totalPages);
    expect(AgentDocument.load(agent.save()).inspect({ kind: "overview" })).toMatchObject({ sections: 1 });
  });

  it("exposes all canonical edit kinds through progressive capabilities", () => {
    const agent = AgentDocument.create();
    const capabilities = agent.capabilities();
    expect(capabilities.map((capability) => capability.kind)).toEqual(INTENT_KINDS);
    expect(new Set(capabilities.map((capability) => capability.kind)).size).toBe(capabilities.length);
    for (const capability of capabilities) expect(capability.inputSchema.additionalProperties).toBe(false);
    expect(validateAgentOperationShape({ kind: "acceptAllRevisions" })).toBeNull();
    expect(validateAgentOperationShape({ kind: "acceptAllRevisions", extra: true })).toContain("does not allow extra");
    expect(validateAgentOperationShape({ kind: "setPageLayout", patch: { pleadingPaper: true } })).toContain("does not allow pleadingPaper");
    expect(validateAgentOperationShape({
      kind: "insertChart",
      runRef: "run:1",
      chart: { type: "column", categories: ["A"], series: [{ name: "S", values: [1], extra: true }] },
    })).toContain("does not allow extra");
    expect(agent.tools().map((tool) => tool.name)).toEqual([
      "word_document_capabilities",
      "word_document_compose",
      "word_document_inspect",
      "word_document_edit",
      "word_document_project",
      "word_document_patch",
      "word_document_asset",
      "word_document_save",
    ]);
    const projectSchema = agent.tools().find((tool) => tool.name === "word_document_project")?.inputSchema;
    expect(projectSchema).toMatchObject({ type: "object", additionalProperties: false });
    const patchSchema = agent.tools().find((tool) => tool.name === "word_document_patch")?.inputSchema;
    expect(patchSchema).toMatchObject({ type: "object", additionalProperties: false, required: ["revision"] });
    const composeSchema = agent.tools().find((tool) => tool.name === "word_document_compose")?.inputSchema;
    expect(composeSchema).toMatchObject({ type: "object", additionalProperties: false });
    const inspectSchema = agent.tools().find((tool) => tool.name === "word_document_inspect")?.inputSchema;
    expect(inspectSchema).toMatchObject({ type: "object", additionalProperties: false, required: ["kind"] });
    expect((inspectSchema?.properties as Record<string, Record<string, unknown>>).kind.enum).toHaveLength(6);
    const editSchema = agent.tools().find((tool) => tool.name === "word_document_edit")?.inputSchema;
    const operationSchemas = (((editSchema?.properties as Record<string, unknown>).operations as Record<string, unknown>).items as Record<string, unknown>).anyOf;
    expect(operationSchemas).toHaveLength(INTENT_KINDS.length);
    // The Messages API rejects anyOf/oneOf/allOf at the TOP level of
    // input_schema (nested, like the operations items above, is fine). Every
    // advertised tool must satisfy that, or the first real request 400s.
    for (const tool of agent.tools()) {
      expect(tool.inputSchema.type, `${tool.name} root type`).toBe("object");
      for (const key of ["anyOf", "oneOf", "allOf"]) {
        expect(tool.inputSchema[key], `${tool.name} top-level ${key}`).toBeUndefined();
      }
    }
  });

  it("edits the second equation in a paragraph by index", async () => {
    // A paragraph can hold several equations, and inspection shows every one.
    // mathIndex is what lets an agent edit the one it just read instead of
    // always landing on the first.
    const agent = AgentDocument.create();
    const target = () => {
      const read = agent.inspect({ kind: "read" });
      if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
      const paragraph = read.blocks[0];
      const run = paragraph.runs.find((candidate) => /^run:\d+$/.test(candidate.ref));
      if (!run) throw new Error("no editable run");
      return { blockRef: paragraph.ref, runRef: run.ref };
    };
    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertText", at: { ...target(), offset: 0 }, text: "Both hold:" }],
    });
    for (const mathText of ["c+d", "a+b"]) {
      await agent.edit({ revision: agent.revision, operations: [{ kind: "insertMath", runRef: target().runRef, mathText }] });
    }
    const { blockRef } = target();
    const equationsOf = () => DocxDocument.load(agent.save()).pkg.text("word/document.xml").split("<m:oMath>").slice(1);
    expect(equationsOf()).toHaveLength(2);

    // Each insert lands ahead of the last, so document order is "a+b" then
    // "c+d" — index 1 names "c+d".
    await agent.edit({ revision: agent.revision, operations: [{ kind: "setMathLinear", blockRef, mathIndex: 1, mathText: "z^2" }] });
    const equations = equationsOf();
    expect(equations[0]).toContain("<m:t>a+b</m:t>");
    expect(equations[1]).toContain("<m:t>z</m:t>");
    expect(equations[1]).not.toContain("c+d");

    // Omitting the index still means the first equation.
    await agent.edit({ revision: agent.revision, operations: [{ kind: "setMathLinear", blockRef, mathText: "q" }] });
    expect(equationsOf()[0]).toContain("<m:t>q</m:t>");

    // The schema takes the field and rejects a negative one.
    expect(validateAgentOperationShape({ kind: "deleteMath", blockRef: "block:1", mathIndex: 2 })).toBeNull();
    expect(validateAgentOperationShape({ kind: "deleteMath", blockRef: "block:1", mathIndex: -1 })).toContain("minimum");
  });

  it("exposes insertMergeField and insertCitation, and applies both", async () => {
    // Capability rows come straight from the core registry declaration.
    const capabilities = AgentDocument.create().capabilities();
    const byKind = new Map(capabilities.map((capability) => [capability.kind, capability]));
    expect(byKind.get("insertMergeField")).toMatchObject({ category: "insert", required: ["runRef", "name"] });
    expect(byKind.get("insertCitation")).toMatchObject({ category: "insert", required: ["runRef", "tag"] });
    expect(validateAgentOperationShape({ kind: "insertCitation", runRef: "run:1", tag: "Doe03" })).toBeNull();
    expect(validateAgentOperationShape({ kind: "insertCitation", runRef: "run:1", tag: "no spaces" })).not.toBeNull();

    const agent = AgentDocument.load(
      zipSync({
        "[Content_Types].xml": strToU8(
          `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
            `<Default Extension="xml" ContentType="application/xml"/>` +
            `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
            `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        ),
        "_rels/.rels": strToU8(
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
        ),
        "word/_rels/document.xml.rels": strToU8(
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`,
        ),
        "customXml/item1.xml": strToU8(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA">` +
            `<b:Source><b:Tag>Doe03</b:Tag>` +
            `<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last></b:Person></b:NameList></b:Author></b:Author>` +
            `<b:Title>A Study of Things</b:Title><b:Year>2003</b:Year></b:Source></b:Sources>`,
        ),
        "word/document.xml": strToU8(body(`<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`)),
      }),
    );
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const runRef = read.blocks[0].runs[0].ref;
    const result = await agent.edit({
      revision: agent.revision,
      operations: [
        { kind: "insertMergeField", runRef, name: "First Name" },
        { kind: "insertCitation", runRef, tag: "Doe03" },
      ],
    });
    expect(result.status).toBe("applied");
    const saved = DocxDocument.load(agent.save());
    expect(textOf(saved)).toContain("«First Name»");
    expect(textOf(saved)).toContain("(Doe, 2003)");
  });
});
