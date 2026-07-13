import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { applyRunFormat, SelectionSegment } from "../src/edit/commands.js";
import { addComment } from "../src/edit/comments.js";
import { setListType, setListLevel } from "../src/edit/lists.js";
import { setLink, removeLink, linkAt } from "../src/edit/links.js";
import { adjustIndent, setParagraphSpacing } from "../src/edit/paragraph.js";
import { findAll, replaceAll, transformCase } from "../src/edit/find.js";
import { applyTableOp } from "../src/edit/tables.js";
import { imageAltText, setImageAltText, replaceImageBlip } from "../src/edit/images.js";
import { insertFootnote } from "../src/edit/notes.js";
import { insertPageField } from "../src/edit/fields.js";
import { linearizeMath, parseMathLinear, setMathLinear, mathLinearOf, isLinearSafe } from "../src/edit/math.js";
import { XmlElement } from "../src/xml.js";
import { serializeXml, parseXml } from "../src/xml.js";
import { makeDocx, makeDocxWithMedia, wrapDocument, p } from "./helpers.js";
import { Paragraph, Run, TextContent } from "../src/model.js";

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
    applyRunFormat(doc, [segFor(run, 6, 11)], { bold: true });
    expect((firstRun(doc).para.children as Run[]).length).toBe(3);
    history.undo();
    expect((firstRun(doc).para.children as Run[]).length).toBe(1);
    expect(textOf(firstRun(doc).para)).toBe("Hello brave world");
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
    expect(applyTableOp(doc, tOf(doc, "B2"), { kind: "cellShading", fill: "FFF2CC" })).toBe(true);
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
    const cases = ["e^x=1+x+x/2", "a_i+b^{2y}", "{a+b}/{2c}", "√{x+1}"];
    for (const c of cases) {
      expect(linearizeMath(parseMathLinear(c))).toBe(c);
    }
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
