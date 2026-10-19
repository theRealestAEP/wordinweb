import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import type { TextItem } from "../src/layout/types.js";
import { FontSet } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, wrapDocument } from "./helpers.js";

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

/** A 260 tw (17.33 px) exact row holding three 16 px lines: ROWA fits, ROWB
 * starts inside the box and crosses its bottom, ROWC starts below it. */
function exactRowDocx(): Uint8Array {
  const LINE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>';
  const para = (text: string) => `<w:p><w:pPr>${LINE}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const body =
    para("TOP") +
    `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="260"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${para("ROWA")}${para("ROWB")}${para("ROWC")}</w:tc></w:tr></w:tbl>` +
    para("MARK") +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  return makeDocx({ "word/document.xml": wrapDocument(body) });
}

async function render(docx: Uint8Array): Promise<Buffer> {
  const chunks: Buffer[] = [];
  const out = new Writable({
    write(chunk: Buffer, _encoding, callback) {
      chunks.push(chunk);
      callback();
    },
  });
  await renderDocxToPdf(bytesSource(docx), out);
  return Buffer.concat(chunks);
}

function contentStreams(bytes: Buffer): string {
  const text = bytes.toString("latin1");
  const streams = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)];
  return streams.map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1")).join("\n");
}

describe("a line an exact row clips", () => {
  it("is laid with the row box as its clip; a line below the box is dropped", () => {
    const doc = DocxDocument.load(exactRowDocx());
    const { pages } = layoutDocument(doc, { measurer: new NodeMeasurer(new FontSet()) });
    const items = pages[0].items.filter((item): item is TextItem => item.kind === "text");
    const text = (want: string) => items.find((item) => item.text === want);
    expect(text("ROWA")?.clip).toBeUndefined();
    const rowB = text("ROWB")!;
    expect(rowB.clip).toBeDefined();
    expect(rowB.clip!.bottom - rowB.clip!.top).toBeCloseTo(260 / 15, 1);
    expect(rowB.lineTop).toBeLessThan(rowB.clip!.bottom);
    expect(rowB.lineTop + rowB.lineHeight).toBeGreaterThan(rowB.clip!.bottom);
    expect(text("ROWC")).toBeUndefined();
  });

  it("is drawn inside a clip rectangle of the row box, in the text layer as Word's PDFs have it", async () => {
    const bytes = await render(exactRowDocx());
    const content = contentStreams(bytes);
    // One clip: the row box, 13 pt tall (260 tw), before ROWB's glyphs and after ROWA's.
    const clips = [...content.matchAll(/q\n([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re W n\n/g)];
    expect(clips).toHaveLength(1);
    expect(Number(clips[0][4])).toBeCloseTo(13, 1);
    const glyphs = [...content.matchAll(/\] TJ|> Tj/g)].map((m) => m.index!);
    // Four texts are drawn (TOP, ROWA, ROWB, MARK): the clip sits before the third.
    expect(glyphs).toHaveLength(4);
    expect(clips[0].index!).toBeGreaterThan(glyphs[1]);
    expect(clips[0].index!).toBeLessThan(glyphs[2]);
    if (!hasPdftotext) return;
    const dir = mkdtempSync(join(tmpdir(), "wordinweb-pdf-"));
    const path = join(dir, "doc.pdf");
    writeFileSync(path, bytes);
    const text = execFileSync("pdftotext", [path, "-"], { encoding: "utf8" });
    expect(text).toContain("ROWA");
    expect(text).toContain("ROWB");
    expect(text).not.toContain("ROWC");
  });
});
