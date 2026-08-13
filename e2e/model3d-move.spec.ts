import { test, expect, type Page } from "@playwright/test";
import { deflateSync } from "node:zlib";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE, goLive } from "./_helpers";

/**
 * USER-REPORTED (#153): a 3D model's "Move" button drags it to the WRONG
 * place, then stops moving it at all. Measured before the fix, three +60/0
 * drags gave +138.8/+389.3, +138.8/+389.3, then 0/0 — the object never lands
 * where the pointer put it, and it walks vertically from a drag with no
 * vertical component.
 *
 * The claim under test is not "it moved". It is that the object's landing
 * spot equals the drag, three drags running. A floating PICTURE rides along
 * as the control: it takes the same move path and already lands exactly, so
 * a fix that quietly breaks it fails here rather than in a user's document.
 *
 * Real mouse input on purpose: dispatchEvent hits an element whether or not
 * the browser could ever route a click there, and the Move button sits on a
 * selection overlay stacked over a live <model-viewer>.
 */

/** A real, decodable PNG — the poster is decoded by the browser. */
function makePng(width: number, height: number): Buffer {
  const crcTable: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  const crc32 = (buf: Buffer): number => {
    let c = 0xffffffff;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.alloc(height * (1 + width * 3));
  for (let y = 0; y < height; y++) {
    const row = y * (1 + width * 3);
    raw[row] = 0;
    for (let x = 0; x < width; x++) {
      const p = row + 1 + x * 3;
      raw[p] = (x * 255) / width;
      raw[p + 1] = (y * 255) / height;
      raw[p + 2] = 128;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const EMU_PER_PX = 9525;
const px = (n: number) => String(Math.round(n * EMU_PER_PX));

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"',
  'mc:Ignorable="w14 wp14"',
].join(" ");

/**
 * A wp:anchor around `graphic`. `hRel` is the horizontal frame: Word anchors
 * a 3D model to the COLUMN and a dropped picture to the MARGIN, and the two
 * origins differ — which is exactly the difference the move path has to
 * absorb, so the fixture keeps them different.
 */
function anchor(graphic: string, opts: { hRel: string; x: number; y: number; w: number; h: number; wrap: string }): string {
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${opts.hRel}"><wp:posOffset>${px(opts.x)}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${px(opts.y)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${px(opts.w)}" cy="${px(opts.h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `${opts.wrap}` +
    `<wp:docPr id="${opts.hRel === "column" ? 11 : 12}" name="Object ${opts.hRel}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `${graphic}` +
    `</wp:anchor></w:drawing>`
  );
}

const MODEL_GRAPHIC =
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">` +
  `<am3d:model3d r:embed="rId5">` +
  `<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${px(144)}" cy="${px(144)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>` +
  `<am3d:raster><am3d:blip r:embed="rId4"/></am3d:raster>` +
  `</am3d:model3d></a:graphicData></a:graphic>`;

const PIC_GRAPHIC =
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="control.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${px(96)}" cy="${px(96)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic>`;

/**
 * One page holding a floating 3D MODEL in an indented sidebar table and a
 * floating PICTURE anchored in ordinary body text above it — the object that
 * reproduces, and the control that does not.
 */
function modelAndPictureDocx(): Buffer {
  const para = (text: string, extra = "") =>
    `<w:p>${extra ? `<w:r>${extra}</w:r>` : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const model = anchor(MODEL_GRAPHIC, { hRel: "column", x: 6, y: 0, w: 144, h: 144, wrap: "<wp:wrapNone/>" });
  // Clear ABOVE and to the right of the sidebar, so a click aimed at one
  // object can never land on the other, before or after three +60px drags.
  const picture = anchor(PIC_GRAPHIC, { hRel: "margin", x: 380, y: 0, w: 96, h: 96, wrap: '<wp:wrapSquare wrapText="bothSides"/>' });
  const filler = Array.from({ length: 10 }, (_, i) => para(`Body line ${i} of the document.`));
  // THE MODEL LIVES IN A TABLE CELL, as Word's cover-letter template puts it
  // (an indented sidebar). That is what makes this document reproduce and a
  // plain body float not: the cell gives the anchor its own origin on BOTH
  // axes — indented from the page margin, and down at the cell's paragraph —
  // and a page-relative write lands that far off on each.
  const sidebar =
    `<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="360" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>` +
    para("Sidebar", model) +
    `</w:tc></w:tr></w:tbl>`;
  const body =
    para("Curriculum vitae") +
    para("Anchor for the control picture.", picture) +
    filler.slice(0, 3).join("") +
    sidebar +
    filler.slice(3).join("") +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Default Extension="glb" ContentType="model/gltf-binary"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
        `<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d" Target="media/model3d1.glb"/>` +
        `</Relationships>`,
    ),
    "word/media/image1.png": new Uint8Array(makePng(64, 64)),
    // A minimal glTF-binary header: enough for the renderer to mount a viewer.
    "word/media/model3d1.glb": new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]),
  };
  return Buffer.from(zipSync(files));
}

/**
 * Open the fixture through the demo's File ▸ Open .docx input.
 *
 * `live` picks the ROUTE, and both have to be exercised. The demo's local
 * editor sets no `onIntent`, so it takes the local branch; the desktop app
 * pins its local editing onto the collab route (engine 9166144), so its
 * `onIntent` is defined even offline and every edit takes the intent branch.
 * The two write positions through different code, and a fix proved on one
 * says nothing about the other.
 */
async function openFixture(page: Page, live = true): Promise<void> {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*=".docx"]').setInputFiles({
    name: "model-and-picture.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: modelAndPictureDocx(),
  });
  await expect(page.locator(PAGE).first()).toBeVisible();
  await expect(page.locator("[data-dxw-model3d]")).toHaveCount(1);
  await expect(page.locator(`${PAGE} img`)).toHaveCount(1);
  if (!live) return;
  await goLive(page);
  await expect(page.locator("[data-dxw-model3d]")).toHaveCount(1);
}

/** Select an object, then put it at an exact page position through the wrap
 * bar's Position command — the same write a drag makes, with no gesture. */
async function positionAt(page: Page, selector: string, x: number, y: number): Promise<void> {
  await inView(page, selector);
  const box = (await page.locator(selector).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + 20);
  await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);
  // The object's OWN wrap bar, not the ribbon's same-named button.
  await page.locator('button[title="Exact page position (px)"]').click();
  const dialog = page.locator("[data-dxw-number-pair-dialog]");
  await expect(dialog).toHaveCount(1);
  await dialog.getByLabel("X (pixels)").fill(String(x));
  await dialog.getByLabel("Y (pixels)").fill(String(y));
  await dialog.getByRole("button", { name: "Apply" }).click();
  await expect(dialog).toHaveCount(0);
  await page.waitForTimeout(350);
}

/** Press-move-release with a real mouse, exactly `dx`/`dy` client px. */
async function dragBy(page: Page, from: { x: number; y: number }, dx: number, dy: number): Promise<void> {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(from.x + dx / 2, from.y + dy / 2);
  await page.mouse.move(from.x + dx, from.y + dy);
  await page.mouse.up();
  await page.waitForTimeout(350);
}

const centre = (b: { x: number; y: number; width: number; height: number }) => ({
  x: b.x + b.width / 2,
  y: b.y + b.height / 2,
});

/**
 * The object's position in DOCUMENT px (its inline left/top on the page
 * surface). Client rects would answer a different question here: a drag can
 * carry the object below the fold, and every later measurement would then be
 * reporting the scroll as movement.
 */
async function at(page: Page, selector: string): Promise<{ left: number; top: number }> {
  return page.evaluate((sel) => {
    const el = document.querySelector<HTMLElement>(sel)!;
    return { left: parseFloat(el.style.left) || 0, top: parseFloat(el.style.top) || 0 };
  }, selector);
}

/** Scroll the object into view and hand back a grab point that is really on
 * screen — a press outside the viewport hits nothing, which reads exactly
 * like "the object stopped responding". */
async function inView(page: Page, selector: string): Promise<void> {
  await page.evaluate((sel) => document.querySelector<HTMLElement>(sel)!.scrollIntoView({ block: "center" }), selector);
  await page.waitForTimeout(250);
}

async function onScreen(page: Page, box: { x: number; y: number }): Promise<boolean> {
  const vp = page.viewportSize()!;
  return box.x > 0 && box.y > 0 && box.x < vp.width && box.y < vp.height;
}

/**
 * The same defect with no gesture at all, and with an exact expected answer
 * rather than a delta. "Exact page position" reads the object's own CSS
 * position — page-surface px — and writes the number straight back through
 * setFloatingPagePosition, which resolves it against the ANCHOR's origin
 * instead. Ask for (0,0) and the object lands ON its anchor origin, which
 * names the defect: landed = requested + anchorOrigin.
 *
 * Both routes, because they are different code. Both were wrong.
 */
for (const live of [false, true]) {
  const route = live ? "collab route (what the desktop app runs)" : "local route";
  test.describe(`#153 exact page position, ${route}`, () => {
    test("the model lands on the page position it was given", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, "[data-dxw-model3d]", 0, 0);
      const got = await at(page, "[data-dxw-model3d]");
      expect(
        Math.abs(got.left) < 1 && Math.abs(got.top) < 1,
        `Position (0,0) must land the model at (0,0), got (${got.left.toFixed(3)}, ${got.top.toFixed(3)})`,
      ).toBe(true);
    });

    test("and on a second, non-zero position, without compounding", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, "[data-dxw-model3d]", 100, 200);
      const first = await at(page, "[data-dxw-model3d]");
      await positionAt(page, "[data-dxw-model3d]", 140, 240);
      const second = await at(page, "[data-dxw-model3d]");
      expect(
        Math.abs(first.left - 100) < 1 && Math.abs(first.top - 200) < 1 &&
        Math.abs(second.left - 140) < 1 && Math.abs(second.top - 240) < 1,
        `Position must be absolute: (100,200) gave (${first.left.toFixed(3)}, ${first.top.toFixed(3)}), ` +
        `(140,240) gave (${second.left.toFixed(3)}, ${second.top.toFixed(3)})`,
      ).toBe(true);
    });

    test("the floating picture control keeps landing where it is put", async ({ page }) => {
      await openFixture(page, live);
      await positionAt(page, `${PAGE} img`, 120, 60);
      const got = await at(page, `${PAGE} img`);
      expect(
        Math.abs(got.left - 120) < 1 && Math.abs(got.top - 60) < 1,
        `Position (120,60) must land the picture at (120,60), got (${got.left.toFixed(3)}, ${got.top.toFixed(3)})`,
      ).toBe(true);
    });
  });
}

test.describe("#153 a 3D model's Move button", () => {
  test("three +60px drags leave the model exactly +180px across and 0 down", async ({ page }) => {
    await openFixture(page);
    const start = await at(page, "[data-dxw-model3d]");

    for (let i = 1; i <= 3; i++) {
      await inView(page, "[data-dxw-model3d]");
      // Select the object: its Move button only exists on the selection
      // overlay. Press above the centre rotate puck, which owns the middle.
      const box = (await page.locator("[data-dxw-model3d]").boundingBox())!;
      await page.mouse.click(box.x + box.width / 2, box.y + 20);
      const move = page.locator("[data-dxw-object-move]");
      await expect(move, `the Move button should be on the overlay before drag ${i}`).toHaveCount(1);
      const grab = centre((await move.boundingBox())!);
      expect(await onScreen(page, grab), `the Move button must be on screen for drag ${i}`).toBe(true);
      await dragBy(page, grab, 60, 0);
    }

    const end = await at(page, "[data-dxw-model3d]");
    const dx = end.left - start.left;
    const dy = end.top - start.top;
    expect(
      Math.abs(dx - 180) < 1 && Math.abs(dy) < 1,
      `three +60/0 drags must land the model at +180/0, got +${dx.toFixed(1)}/+${dy.toFixed(1)}`,
    ).toBe(true);
  });

  test("the floating picture control still lands exactly where it is dragged", async ({ page }) => {
    await openFixture(page);
    const start = await at(page, `${PAGE} img`);

    for (let i = 1; i <= 3; i++) {
      await inView(page, `${PAGE} img`);
      const box = (await page.locator(`${PAGE} img`).first().boundingBox())!;
      await page.mouse.click(centre(box).x, centre(box).y);
      await expect(page.locator("[data-dxw-object-selection]")).toHaveCount(1);
      const grab = centre((await page.locator(`${PAGE} img`).first().boundingBox())!);
      expect(await onScreen(page, grab), `the picture must be on screen for drag ${i}`).toBe(true);
      await dragBy(page, grab, 60, 0);
    }

    const end = await at(page, `${PAGE} img`);
    const dx = end.left - start.left;
    const dy = end.top - start.top;
    expect(
      Math.abs(dx - 180) < 1 && Math.abs(dy) < 1,
      `three +60/0 drags must land the picture at +180/0, got +${dx.toFixed(1)}/+${dy.toFixed(1)}`,
    ).toBe(true);
  });
});
