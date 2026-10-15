import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertChartAt } from "../src/edit/charts.js";
import type { ChartData, Paragraph, Run, TextContent } from "../src/model.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** A document whose only chart is the one insertChart just authored, as bytes. */
export function chartDocx(data: ChartData): Uint8Array {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Anchor")) }));
  const paragraph = doc.sections[0].blocks[0] as Paragraph;
  const run = paragraph.children[0] as Run;
  const text = run.content[0] as TextContent;
  insertChartAt(doc, text.srcT!, data);
  return doc.save();
}

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
}

async function render(docx: Uint8Array): Promise<{ bytes: Buffer; content: string }> {
  const out = new Collect();
  await renderDocxToPdf(bytesSource(docx), out);
  const bytes = Buffer.concat(out.chunks);
  const text = bytes.toString("latin1");
  const streams = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)];
  const content = streams
    .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
    .filter((s) => s.includes(" re") || s.includes(" Tf"))
    .join("\n");
  return { bytes, content };
}

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

function pdfText(bytes: Buffer): string {
  const dir = mkdtempSync(join(tmpdir(), "wordinweb-pdf-"));
  const path = join(dir, "doc.pdf");
  writeFileSync(path, bytes);
  return execFileSync("pdftotext", [path, "-"], { encoding: "utf8" }).replace(/\s+/g, " ");
}

describe("renderDocxToPdf charts", () => {
  const column: ChartData = {
    type: "column",
    title: "Architecture Decision Profile",
    categories: ["Control", "Delivery speed", "Operating effort"],
    series: [
      { name: "Managed platform", values: [8, 9, 4] },
      { name: "Custom stack", values: [10, 6, 9] },
    ],
    legend: "b",
  };

  it("draws a column chart's bars as filled rectangles in the palette", async () => {
    const { content } = await render(chartDocx(column));
    // Six bars, each a solid rect in one of the Office accents.
    const accent1 = content.match(/0\.082 0\.376 0\.51 rg [\d. ]+ re f/g)?.length ?? 0;
    const accent2 = content.match(/0\.914 0\.443 0\.196 rg [\d. ]+ re f/g)?.length ?? 0;
    expect(accent1).toBeGreaterThanOrEqual(3);
    expect(accent2).toBeGreaterThanOrEqual(3);
    // Gridlines and axes in Word's rule gray.
    expect(content).toMatch(/0\.851 0\.851 0\.851 RG/);
  });

  it.skipIf(!hasPdftotext)("labels the title, legend and scale as searchable text", async () => {
    const { bytes } = await render(chartDocx(column));
    const text = pdfText(bytes);
    expect(text).toContain("Architecture Decision Profile");
    expect(text).toContain("Managed platform");
    expect(text).toContain("Custom stack");
    for (const tick of ["0", "2", "4", "6", "8", "10", "12"]) expect(text).toContain(` ${tick} `);
    expect(text).toContain("Delivery speed");
  });

  it("draws a pie as wedges of Bézier arcs", async () => {
    const { content } = await render(chartDocx({ type: "pie", categories: ["a", "b", "c"], series: [{ name: "s", values: [1, 2, 3] }] }));
    // Three wedges, each filled and outlined in white.
    expect(content.match(/1 1 1 RG [^Q]* c [^Q]*h B Q/g)?.length).toBe(3);
  });
});
