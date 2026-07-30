import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING } from "./_helpers";

/**
 * BIG-DOCUMENT MEMORY REGRESSION — the "500-page grant document eats the
 * tab" report: heap growing at tens of MB/s and an editor that swallowed
 * keystrokes.
 *
 * Root cause pinned here: typing into a BOOKMARKED paragraph (every heading
 * a TOC targets) failed reparseBodyParagraph, so each keystroke after a
 * caret move ran doc.refresh() + a full async relayout behind an inert
 * container. Wall-time suites never caught it because the keystrokes that
 * DID land were fast — the signal is allocation, so this suite measures
 * heap. (bigdoc-typing.spec.ts pins the same flow for latency on plain
 * paragraphs; this one types into a heading, which was the broken path.)
 *
 * Metrics are emitted as STRESS-METRIC lines (internal/perf shape).
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const HEADINGS = 400;
const BODY_PER = 12; // ~218 pages, safely past BACKGROUND_LAYOUT_PAGE_THRESHOLD
const KEYS = 25;

/** A document shaped like the report: a long TOC with dotted leaders and
 * PAGEREF fields, then bookmarked headings each followed by body text. */
function tocDocx(): Buffer {
  let body = `<w:p><w:r><w:t>Table of Contents</w:t></w:r></w:p>`;
  for (let i = 0; i < HEADINGS; i++) {
    body +=
      `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>` +
      `<w:r><w:t xml:space="preserve">Section ${i}: Specific Aims and Research Strategy</w:t></w:r>` +
      `<w:r><w:tab/></w:r>` +
      `<w:fldSimple w:instr=" PAGEREF _Toc${1000 + i} \\h "><w:r><w:t>1</w:t></w:r></w:fldSimple>` +
      `</w:p>`;
  }
  for (let i = 0; i < HEADINGS; i++) {
    body +=
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
      `<w:bookmarkStart w:id="${i}" w:name="_Toc${1000 + i}"/>` +
      `<w:r><w:t xml:space="preserve">HDG${i} Specific Aims and Research Strategy</w:t></w:r>` +
      `<w:bookmarkEnd w:id="${i}"/></w:p>`;
    for (let j = 0; j < BODY_PER; j++) {
      body += `<w:p><w:r><w:t xml:space="preserve">Body ${i}.${j}: the quick brown fox jumps over the lazy dog while the committee deliberates at length about the research strategy. </w:t></w:r></w:p>`;
    }
  }
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return Buffer.from(zipSync({
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
  }));
}

function metric(scenario: string, fields: Record<string, string | number>): void {
  const body = Object.entries(fields)
    .map(([k, v]) => `${k}=${typeof v === "number" ? (Number.isInteger(v) ? v : v.toFixed(1)) : v}`)
    .join(" ");
  console.log(`STRESS-METRIC ${scenario} ${body}`);
}

async function layoutBusy(page: Page): Promise<boolean> {
  return page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]"));
}

test("typing into a bookmarked heading stays incremental (heap + keystrokes)", async ({ page }) => {
  test.setTimeout(300_000);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Performance.enable");
  const heapMB = async () => {
    const { metrics } = await cdp.send("Performance.getMetrics");
    return (metrics.find((m) => m.name === "JSHeapUsedSize")?.value ?? 0) / 1e6;
  };

  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
    name: "toc.docx", mimeType: DOCX_MIME, buffer: tocDocx(),
  });
  await expect(page.locator(".dxw-page").first()).toBeVisible({ timeout: 120_000 });
  await expect.poll(() => layoutBusy(page), { timeout: 60_000 }).toBe(false);
  await page.waitForTimeout(2000);

  // Scroll a heading into view (they live past the TOC pages) and click it —
  // the caret JUMP into a bookmarked paragraph is the regression trigger.
  const first = await page.locator(".dxw-page").first().boundingBox();
  await page.mouse.move(first!.x + first!.width / 2, first!.y + 100);
  for (let i = 0; i < 40; i++) {
    await page.mouse.wheel(0, 2000);
    await page.waitForTimeout(50);
  }
  await page.waitForTimeout(1500);
  const target = await page.evaluate(() => {
    for (const leaf of Array.from(document.querySelectorAll<HTMLElement>(".dxw-page span"))) {
      const t = leaf.textContent ?? "";
      if (!t.includes("HDG")) continue;
      const r = leaf.getBoundingClientRect();
      if (r.width > 0 && r.top > 0 && r.bottom < window.innerHeight) {
        return { x: r.left + 20, y: r.top + r.height / 2 };
      }
    }
    return null;
  });
  expect(target, "a heading must be on screen to type into").toBeTruthy();
  await page.mouse.click(target!.x, target!.y);
  await page.waitForTimeout(300);

  // Type into the heading while sampling heap and the inert state.
  const heapSamples: number[] = [await heapMB()];
  let busySeen = 0;
  const t0 = Date.now();
  for (let i = 0; i < KEYS; i++) {
    await page.keyboard.type("Z");
    await page.waitForTimeout(150);
    if (await layoutBusy(page)) busySeen++;
    heapSamples.push(await heapMB());
  }
  const elapsedS = (Date.now() - t0) / 1000;
  await page.waitForTimeout(3000);
  const landed = await page.evaluate(
    () => (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length,
  );

  const peakMb = Math.max(...heapSamples);
  // Allocation rate = sum of positive heap deltas over elapsed time; net
  // growth alone hides churn that GC keeps reclaiming.
  let alloc = 0;
  for (let i = 1; i < heapSamples.length; i++) {
    const d = heapSamples[i] - heapSamples[i - 1];
    if (d > 0) alloc += d;
  }
  const allocMbPerS = alloc / elapsedS;
  metric("bigdoc-heading-typing", {
    headings: HEADINGS,
    keys: KEYS,
    landed,
    busySeen,
    peakMb,
    allocMbPerS,
  });

  // THE REGRESSION PINS.
  // 1. Every keystroke paints: the pre-fix refresh path left the surface
  //    inert for seconds per keystroke and swallowed input (10/15 landed).
  expect(landed, "keystrokes must not be swallowed").toBe(KEYS);
  // 2. Plain typing must never enter the inert whole-document relayout.
  expect(busySeen, "typing must not trigger the inert relayout").toBe(0);
  // 3. Allocation: the pre-fix path re-parsed the whole model and re-laid
  //    the whole document per keystroke (~90+ MB/s here, 61 MB/s in the
  //    field). The incremental path measures ~15 MB/s; the bound leaves CI
  //    headroom while staying far below the broken figure.
  expect(allocMbPerS, "per-keystroke allocation must stay incremental").toBeLessThan(45);
  // 4. Peak heap: the broken path held several full model generations at
  //    once (measured ~1 GB+ here; ~2.1 GB in the field). Post-fix peak is
  //    ~250 MB on this document.
  expect(peakMb, "peak heap on a 200-page document").toBeLessThan(700);
});
