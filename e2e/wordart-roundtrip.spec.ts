import { expect, test, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import { LANDING, PAGE, clickTextStart } from "./_helpers";

async function insertWordArt(page: Page, text: string, preset: "Plain" | "Arch up" | "Wave"): Promise<void> {
  await page.getByRole("button", { name: "A", exact: true }).click();
  await page.getByLabel("WordArt text").fill(text);
  await page.getByTitle(`Insert WordArt ${preset}`).click();
}

async function dragDrawing(page: Page, index: number, dx: number, dy: number): Promise<void> {
  const drawing = page.locator("[data-dxw-drawing]").nth(index);
  const before = await drawing.boundingBox();
  if (!before) throw new Error(`WordArt ${index} is not visible`);
  const x = before.x + before.width / 2;
  const y = before.y + before.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx, y + dy, { steps: 8 });
  await page.mouse.up();
  await expect.poll(async () => {
    const after = await drawing.boundingBox();
    return after ? Math.hypot(after.x - before.x, after.y - before.y) : 0;
  }).toBeGreaterThan(5);
}

const drawingPositions = (page: Page) =>
  page.locator("[data-dxw-drawing]").evaluateAll((drawings) => drawings.map((drawing) => ({
    x: parseFloat((drawing as HTMLElement).style.left),
    y: parseFloat((drawing as HTMLElement).style.top),
    width: parseFloat((drawing as HTMLElement).style.width),
    height: parseFloat((drawing as HTMLElement).style.height),
  })));

const bodyPositions = (page: Page) =>
  page.locator(`${PAGE} [data-dxw-item-kind="text"]`).evaluateAll((nodes) =>
    nodes
      .filter((node) => ["Alpha", "Bravo", "Charlie", "Delta"].includes(node.textContent ?? ""))
      .map((node) => {
        const rect = (node as HTMLElement).getBoundingClientRect();
        return { text: node.textContent, x: rect.x, y: rect.y };
      }),
  );

test("several WordArt objects drag repeatedly and keep native positions through DOCX reload", async ({ page }) => {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await clickTextStart(page);
  for (const [index, text] of ["Alpha body line", "Bravo body line", "Charlie body line", "Delta body line"].entries()) {
    if (index) await page.keyboard.press("Enter");
    await page.keyboard.type(text);
  }
  const bodyBeforeWordArt = await bodyPositions(page);
  expect(bodyBeforeWordArt).toHaveLength(4);

  await page.getByRole("button", { name: "insert", exact: true }).click();
  await insertWordArt(page, "FIRST ART", "Plain");
  await insertWordArt(page, "SECOND ART", "Arch up");
  await insertWordArt(page, "THIRD ART", "Wave");
  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(3);

  // Decorative text must not carve the body into several wrapping columns.
  // This was the visible failure as soon as two square-wrapped objects shared
  // the same insertion point.
  const bodyAfterWordArt = await bodyPositions(page);
  expect(bodyAfterWordArt).toEqual(bodyBeforeWordArt);

  // The first object moves twice. The other two then move out from the shared
  // insertion point. This exercises the real drag path and its rerenders.
  // Equal-z drawings paint in DOM order, so the last one is the object the
  // pointer can reach while all three still share the insertion point.
  await dragDrawing(page, 2, 150, -45);
  await dragDrawing(page, 2, 45, 70);
  await dragDrawing(page, 1, 75, 145);
  await dragDrawing(page, 0, 285, 245);
  const beforeReload = await drawingPositions(page);
  expect(new Set(beforeReload.map(({ x, y }) => `${Math.round(x)},${Math.round(y)}`)).size).toBe(3);

  await page.getByTestId("file-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("file-download").click();
  const download = await downloadPromise;
  const path = await download.path();
  if (!path) throw new Error("WordArt DOCX download has no local path");

  const xml = strFromU8(unzipSync(readFileSync(path))["word/document.xml"]);
  const anchors = [...xml.matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)].map((match) => match[0]);
  expect(anchors).toHaveLength(3);
  for (const anchor of anchors) {
    expect(anchor).toContain('behindDoc="0"');
    expect(anchor).toContain("<wp:wrapNone/>");
    expect(anchor).not.toContain("<wp:wrapSquare");
    expect(anchor).toContain('<wp:positionH relativeFrom="page">');
    expect(anchor).toContain('<wp:positionV relativeFrom="page">');
    const outer = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(anchor);
    const transform = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(anchor);
    expect(outer?.slice(1)).toEqual(transform?.slice(1));
  }
  expect(xml.match(/<a:prstTxWarp prst="(?:textNoShape|textArchUp|textWave1)">/g)).toHaveLength(3);
  expect(xml.match(/<w:jc w:val="center"\/>/g)).toHaveLength(3);
  expect(xml.match(/<a:noAutofit\/>/g)).toHaveLength(3);

  const input = page.locator('input[type="file"][accept*=".docx"]');
  await input.setInputFiles(path);
  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(3);
  const afterReload = await drawingPositions(page);
  expect(afterReload).toHaveLength(3);
  for (let index = 0; index < 3; index++) {
    expect(afterReload[index].x).toBeCloseTo(beforeReload[index].x, 0);
    expect(afterReload[index].y).toBeCloseTo(beforeReload[index].y, 0);
    expect(afterReload[index].width).toBeCloseTo(beforeReload[index].width, 0);
    expect(afterReload[index].height).toBeCloseTo(beforeReload[index].height, 0);
  }
});
