import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE } from "./_helpers";
import fs from "node:fs";

/**
 * DIAGNOSTIC (not part of the suite unless WW_TRACE=1): capture a Chrome
 * performance trace over CDP while typing into the big document, and
 * attribute time per keystroke to Layout / RecalcStyle / Paint / Script.
 * Confirms or kills the "whole-document reflow" hypothesis before any fix.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = Number(process.env.WW_BENCH_PARAS ?? 12000);

function bigDocx(paras: number): Buffer {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const zipped = zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${DOCX_MIME}.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  });
  return Buffer.from(zipped);
}

async function layoutBusy(page: Page): Promise<boolean> {
  return page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]"));
}

test.describe("bigdoc trace", () => {
  test.skip(process.env.WW_TRACE !== "1", "diagnostic only; run with WW_TRACE=1");

  test("attribute keystroke time by trace category", async ({ page, browser }) => {
    test.setTimeout(300_000);
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
      name: "big.docx",
      mimeType: DOCX_MIME,
      buffer: bigDocx(PARAS),
    });
    await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 60_000 });
    await expect
      .poll(() => page.locator(PAGE).count(), { timeout: 120_000 })
      .toBeGreaterThan(50);
    await expect.poll(() => layoutBusy(page), { timeout: 120_000 }).toBe(false);

    // CPU profile across the typing rounds: names the scripted cost that the
    // timeline only shows as FunctionCall.
    const cdp = await page.context().newCDPSession(page);
    await cdp.send("Profiler.enable");
    await cdp.send("Profiler.setSamplingInterval", { interval: 100 });
    await cdp.send("Profiler.start");

    const tracePath = process.env.WW_TRACE_PATH ?? "/tmp/bigdoc-trace.json";
    await browser.startTracing(page, {
      path: tracePath,
      categories: [
        "devtools.timeline",
        "disabled-by-default-devtools.timeline",
        "disabled-by-default-devtools.timeline.invalidationTracking",
        "blink.user_timing",
      ],
    });

    const box = await page.locator(PAGE).first().boundingBox();
    for (let i = 0; i < 15; i++) {
      await page.mouse.click(box!.x + 30, box!.y + 25 + ((i * 37) % 300));
      await page.keyboard.type("Z");
      await page.waitForTimeout(120);
    }

    await browser.stopTracing();

    const { profile } = (await cdp.send("Profiler.stop")) as {
      profile: {
        nodes: { id: number; callFrame: { functionName: string; url: string; lineNumber: number }; hitCount?: number }[];
        samples?: number[];
        timeDeltas?: number[];
      };
    };
    const nodeById = new Map(profile.nodes.map((n) => [n.id, n]));
    const selfUs = new Map<number, number>();
    const samples = profile.samples ?? [];
    const deltas = profile.timeDeltas ?? [];
    for (let i = 0; i < samples.length; i++) {
      selfUs.set(samples[i], (selfUs.get(samples[i]) ?? 0) + (deltas[i] ?? 0));
    }
    const rows2 = [...selfUs.entries()]
      .map(([id, us]) => ({ n: nodeById.get(id)!, ms: us / 1000 }))
      .sort((a, b) => b.ms - a.ms)
      .slice(0, 30);
    for (const { n, ms } of rows2) {
      const f = n.callFrame;
      const file = f.url.split("/").pop() || f.url;
      console.log(`PROFILE ${f.functionName || "(anonymous)"} @${file}:${f.lineNumber} selfMs=${ms.toFixed(1)}`);
    }

    const trace = JSON.parse(fs.readFileSync(tracePath, "utf8")) as {
      traceEvents: { name: string; ph: string; dur?: number; ts: number; args?: Record<string, unknown> }[];
    };
    const byName = new Map<string, { totalMs: number; count: number; maxMs: number }>();
    for (const ev of trace.traceEvents) {
      if (ev.ph !== "X" || !ev.dur) continue;
      const entry = byName.get(ev.name) ?? { totalMs: 0, count: 0, maxMs: 0 };
      entry.totalMs += ev.dur / 1000;
      entry.count++;
      entry.maxMs = Math.max(entry.maxMs, ev.dur / 1000);
      byName.set(ev.name, entry);
    }
    const rows = [...byName.entries()]
      .sort((a, b) => b[1].totalMs - a[1].totalMs)
      .slice(0, 25);
    for (const [name, e] of rows) {
      console.log(
        `TRACE-CAT ${name} total=${e.totalMs.toFixed(1)}ms count=${e.count} max=${e.maxMs.toFixed(1)}ms`,
      );
    }
    // Forced reflow inside script shows as Layout events nested in FunctionCall
    // (stackTrace present in beginData) — count Layout events specifically.
    const layouts = trace.traceEvents.filter((ev) => ev.ph === "X" && ev.name === "Layout" && ev.dur);
    const forced = layouts.filter((ev) => {
      const beginData = (ev.args as { beginData?: { stackTrace?: unknown[] } } | undefined)?.beginData;
      return Array.isArray(beginData?.stackTrace) && beginData!.stackTrace!.length > 0;
    });
    const sum = (list: typeof layouts) => list.reduce((s, ev) => s + (ev.dur ?? 0) / 1000, 0);
    console.log(
      `TRACE-LAYOUT all=${layouts.length} totalMs=${sum(layouts).toFixed(1)} forced=${forced.length} forcedMs=${sum(forced).toFixed(1)}`,
    );
  });
});
