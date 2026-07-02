import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { applyRunFormat, SelectionSegment } from "../src/edit/commands.js";
import { serializeXml, parseXml } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
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
